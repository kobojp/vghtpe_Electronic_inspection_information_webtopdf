"""Isolated real API for browser tests; never requests the hospital website."""
import argparse
import shutil
import os
import secrets
import sys
import tempfile
from pathlib import Path
from unittest import mock

import pymupdf
import uvicorn
from fastapi import HTTPException, Request

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from backend.app import create_app
from desktop.drop import dropped_pdf_files
from backend.config import save_settings
from backend.reports import valid_pdf
from backend.tasks import TaskCancelled, TaskManager


def make_pdf(path, text):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with pymupdf.open() as document:
        document.new_page().insert_text((72, 72), text)
        document.save(path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=18765)
    args = parser.parse_args()
    root = Path(__file__).resolve().parent.parent
    cache = root / ".cache"
    cache.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ui-test-", dir=cache) as temporary:
        directory = Path(temporary)
        source = directory / "input" / "2026-01" / "sample.pdf"
        make_pdf(source, "TARGET January")
        second = directory / "input" / "2026-02" / "sample.pdf"
        make_pdf(second, "TARGET February")
        text_file = directory / "note.txt"
        text_file.write_text("text", encoding="utf-8")
        corrupt = directory / "broken.pdf"
        corrupt.write_bytes(b"not a PDF")
        settings_path = directory / "settings.json"
        save_settings({"output_folder": str(directory / "reports"), "extract_folder": str(directory / "extracts"), "open_folder_after_completion": False}, settings_path)

        class Dialogs:
            def select_pdfs(self):
                return [str(source)]
            def select_folder(self):
                return str(directory / "chosen-output")

        def download(downloader, url, destination, name, max_retries=3, log=None):
            if valid_pdf(destination):
                return "skipped"
            if downloader.cancel_event.wait(1.5):
                raise TaskCancelled()
            make_pdf(destination, "Test download")
            return "downloaded"

        data_path = directory / "data.json"
        shutil.copyfile(root / "data.json", data_path)
        app = create_app(data_path=data_path, settings_path=settings_path, dialogs=Dialogs())
        token = os.environ.get("VGHTPE_UI_TEST_TOKEN") or secrets.token_urlsafe(32)
        server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=args.port, log_level="warning"))

        def authenticate(request):
            if not secrets.compare_digest(request.headers.get("X-Test-Token", ""), token):
                raise HTTPException(403)

        @app.get("/__test__/health")
        def test_health(request: Request):
            authenticate(request)
            return {"ready": True}

        @app.get("/__test__/drop-files")
        def test_drop_files(request: Request):
            authenticate(request)
            return dropped_pdf_files({"dataTransfer": {"files": [
                {"name": path.name, "pywebviewFullPath": str(path)} for path in (source, second, text_file, corrupt)
            ]}})

        @app.post("/__test__/reset")
        def test_reset(request: Request):
            authenticate(request)
            app.state.runtime.tasks.shutdown()
            app.state.runtime.tasks = TaskManager()
            shutil.copyfile(root / "data.json", data_path)
            return {"reset": True}

        @app.post("/__test__/shutdown")
        def test_shutdown(request: Request):
            authenticate(request)
            server.should_exit = True
            return {"stopping": True}

        static_route = next(route for route in app.router.routes if getattr(route, "name", None) == "frontend")
        app.router.routes.remove(static_route)
        app.router.routes.append(static_route)
        with mock.patch("backend.app.ReportDownloader.download_report", autospec=True, side_effect=download), mock.patch("backend.app.config.open_folder"):
            server.run()


if __name__ == "__main__":
    main()
