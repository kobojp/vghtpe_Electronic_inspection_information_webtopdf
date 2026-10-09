"""Local HTTP API shared by the React UI and desktop launcher."""
import datetime
import logging
import secrets
import sys
import threading
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from backend import config
from backend.catalog import CatalogConflict, CatalogStore
from backend.pdf_core import extract_pdf
from backend.reports import REPORT_TYPE_MAP, ReportDownloader, get_last_month, get_month_range, build_report_url
from backend.tasks import TaskBusy, TaskCancelled, TaskManager

logger = logging.getLogger("vghtpe")


class DownloadRequest(BaseModel):
    report_ids: list[str] = Field(min_length=1)
    start_month: str
    end_month: str
    catalog_revision: str | None = None


class ExtractRequest(BaseModel):
    paths: list[str] = Field(min_length=1)
    keyword: str = Field(min_length=1, max_length=200)
    keep_contractor_page: bool = True


class SettingsRequest(BaseModel):
    output_folder: str
    extract_folder: str
    open_folder_after_completion: bool


class FolderRequest(BaseModel):
    kind: Literal["download", "extract", "task"]


class ReportFields(BaseModel):
    name: str = Field(min_length=1, max_length=180)
    report_type: str
    api_1: str = Field(min_length=1, max_length=100)
    api_2: str = Field(min_length=1, max_length=100)


class ReportChange(ReportFields):
    revision: str = Field(min_length=1)


class CatalogRevision(BaseModel):
    revision: str = Field(min_length=1)


class Runtime:
    def __init__(self, data_path=None, settings_path=None, dialogs=None):
        self.catalog = CatalogStore(data_path or config.get_data_file_path())
        self.catalog.snapshot()
        self.settings_path = settings_path
        self.settings = config.load_settings(settings_path)
        self.settings_lock = threading.Lock()
        self.tasks = TaskManager()
        self.dialogs = dialogs
        self.token = secrets.token_urlsafe(32)
        self.executable = config.resource_dir() / "wkhtmltopdf.exe"

    def settings_snapshot(self):
        with self.settings_lock:
            return self.settings.copy()

    def finish(self, context, settings, folder):
        context.check_cancelled()
        if settings["open_folder_after_completion"]:
            try:
                config.open_folder(folder)
            except OSError:
                logger.exception("Could not open output folder")
                context.log("任務已處理完成，但無法自動開啟輸出資料夾")


def create_app(data_path=None, settings_path=None, dialogs=None, dev=False):
    runtime = Runtime(data_path, settings_path, dialogs)

    @asynccontextmanager
    async def lifespan(app):
        yield
        runtime.tasks.shutdown()

    app = FastAPI(title="臺北榮總水電消防報表", lifespan=lifespan, docs_url=None, redoc_url=None, openapi_url=None)
    app.state.runtime = runtime

    @app.middleware("http")
    async def local_session(request: Request, call_next):
        if request.url.hostname not in {"127.0.0.1", "localhost", "testserver"}:
            return JSONResponse({"detail": "僅提供本機使用"}, status_code=403)
        origin = request.headers.get("origin")
        allowed = {f"{request.url.scheme}://{request.url.netloc}"}
        if dev:
            allowed.add("http://127.0.0.1:5173")
        if origin and origin not in allowed:
            return JSONResponse({"detail": "不允許此來源"}, status_code=403)
        if request.url.path.startswith("/api/") and request.url.path not in {"/api/health", "/api/bootstrap"}:
            if not secrets.compare_digest(request.headers.get("X-App-Token", ""), runtime.token):
                return JSONResponse({"detail": "連線已失效，請重新開啟程式"}, status_code=403)
        response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        return response

    @app.get("/api/health")
    def health():
        return {"status": "ok"}

    def catalog_snapshot():
        try:
            return runtime.catalog.snapshot()
        except (OSError, ValueError) as error:
            logger.exception("Catalog read failed")
            raise HTTPException(400, f"無法讀取報表資料：{error}") from error

    @app.get("/api/catalog")
    def catalog():
        return {**catalog_snapshot(), "report_types": list(REPORT_TYPE_MAP)}

    def change_catalog(revision, entry=None, report_id=None, delete=False):
        try:
            result = runtime.catalog.mutate(revision, entry, report_id, delete)
        except CatalogConflict as error:
            raise HTTPException(409, str(error)) from error
        except KeyError as error:
            raise HTTPException(404, str(error.args[0])) from error
        except ValueError as error:
            raise HTTPException(422, str(error)) from error
        except OSError as error:
            logger.exception("Catalog write failed")
            raise HTTPException(400, "無法儲存報表資料，請確認 data.json 與所在資料夾可寫入；原資料未替換") from error
        return {**result, "report_types": list(REPORT_TYPE_MAP)}

    @app.post("/api/catalog/reports", status_code=201)
    def add_report(request: ReportChange):
        return change_catalog(request.revision, request.model_dump(exclude={"revision"}))

    @app.put("/api/catalog/reports/{report_id}")
    def edit_report(report_id: str, request: ReportChange):
        return change_catalog(request.revision, request.model_dump(exclude={"revision"}), report_id)

    @app.delete("/api/catalog/reports/{report_id}")
    def delete_report(report_id: str, request: CatalogRevision):
        return change_catalog(request.revision, report_id=report_id, delete=True)

    @app.get("/api/bootstrap")
    def bootstrap():
        catalog = catalog_snapshot()
        return {
            "token": runtime.token,
            "reports": [{key: report[key] for key in ("id", "name", "report_type", "category")} for report in catalog["reports"]],
            "catalog_revision": catalog["revision"],
            "report_types": list(REPORT_TYPE_MAP),
            "last_month": get_last_month(),
            "settings": runtime.settings_snapshot(),
            "native_dialogs": runtime.dialogs is not None,
        }

    @app.get("/api/tasks/current")
    def current_task():
        return runtime.tasks.snapshot()

    @app.post("/api/tasks/{task_id}/cancel")
    def cancel_task(task_id: str):
        try:
            return runtime.tasks.cancel(task_id)
        except KeyError as error:
            raise HTTPException(404, "找不到此任務") from error

    @app.post("/api/tasks/download", status_code=202)
    def download(request: DownloadRequest):
        try:
            months = get_month_range(request.start_month, request.end_month)
        except ValueError as error:
            raise HTTPException(422, str(error)) from error
        catalog = catalog_snapshot()
        if request.catalog_revision is not None and request.catalog_revision != catalog["revision"]:
            raise HTTPException(409, "報表資料已變更，請至報表管理重新載入後再選擇下載")
        ids = set(request.report_ids)
        reports = [report.copy() for report in catalog["reports"] if report["id"] in ids]
        if len(reports) != len(ids):
            raise HTTPException(422, "包含不存在的報表，請重新選擇")
        settings = runtime.settings_snapshot()
        output = Path(settings["output_folder"])

        def work(context):
            downloader = ReportDownloader(runtime.executable, context.cancel_event)
            for month in months:
                for report in reports:
                    context.item(f'{month} / {report["name"]}')
                    path = output / report["category"] / month / f'{report["name"]}.pdf'
                    result = downloader.download_report(build_report_url(report, report["report_type"], month), path, report["name"], log=context.log)
                    context.record(result)
            runtime.finish(context, settings, output)

        try:
            return runtime.tasks.start("download", "報表下載", len(months) * len(reports), output, work)
        except TaskBusy as error:
            raise HTTPException(409, str(error)) from error

    @app.post("/api/tasks/extract", status_code=202)
    def extract(request: ExtractRequest):
        if not request.keyword.strip():
            raise HTTPException(422, "請輸入搜尋關鍵字")
        paths = list(dict.fromkeys(request.paths))
        if any(not Path(path).is_file() or Path(path).suffix.lower() != ".pdf" for path in paths):
            raise HTTPException(422, "請選擇存在的 PDF 檔案")
        settings = runtime.settings_snapshot()
        output = Path(settings["extract_folder"]) / datetime.date.today().strftime("%Y%m%d")

        def work(context):
            for path in paths:
                context.item(Path(path).name)
                try:
                    result = extract_pdf(path, output, request.keyword, request.keep_contractor_page, context)
                except TaskCancelled:
                    raise
                except Exception as error:
                    logger.exception("PDF extraction failed: %s", path)
                    context.log(f"{Path(path).name} 處理失敗：{error}")
                    result = "failed"
                context.record(result)
            runtime.finish(context, settings, output)

        try:
            return runtime.tasks.start("extract", "PDF 擷取", len(paths), output, work)
        except TaskBusy as error:
            raise HTTPException(409, str(error)) from error

    @app.put("/api/settings")
    def update_settings(request: SettingsRequest):
        values = request.model_dump()
        for key in ("output_folder", "extract_folder"):
            if not Path(values[key]).is_absolute():
                raise HTTPException(422, "輸出資料夾須使用完整路徑")
            values[key] = str(Path(values[key]).resolve())
        try:
            with runtime.settings_lock:
                config.save_settings(values, runtime.settings_path)
                runtime.settings = values
        except OSError as error:
            logger.exception("Settings write failed")
            raise HTTPException(400, f"無法儲存設定：{error}") from error
        return values

    @app.post("/api/desktop/select-pdfs")
    def select_pdfs():
        if runtime.dialogs is None:
            raise HTTPException(503, "請透過桌面程式開啟，以使用原生檔案選擇")
        return {"paths": runtime.dialogs.select_pdfs()}

    @app.post("/api/desktop/select-folder")
    def select_folder():
        if runtime.dialogs is None:
            raise HTTPException(503, "請透過桌面程式開啟，以使用原生資料夾選擇")
        return {"path": runtime.dialogs.select_folder()}

    @app.post("/api/desktop/open-folder")
    def open_output(request: FolderRequest):
        settings = runtime.settings_snapshot()
        if request.kind == "task":
            task = runtime.tasks.snapshot()
            if task is None:
                raise HTTPException(404, "尚無任務輸出資料夾")
            folder = task["output_folder"]
        else:
            folder = settings["output_folder" if request.kind == "download" else "extract_folder"]
        try:
            config.open_folder(folder)
        except OSError as error:
            raise HTTPException(400, f"無法開啟資料夾：{error}") from error
        return {"opened": True}

    frontend = config.resource_dir() / ("frontend_dist" if getattr(sys, "frozen", False) else "frontend/dist")
    if frontend.is_dir():
        app.mount("/", StaticFiles(directory=frontend, html=True), name="frontend")
    else:
        @app.get("/")
        def build_required():
            return JSONResponse({"detail": "請先在 frontend 執行 npm run build"}, status_code=503)
    return app
