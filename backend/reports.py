"""Report catalog, month ranges, and atomic HTML-to-PDF downloads."""
import calendar
import datetime
import json
import logging
import os
import subprocess
import tempfile
import threading
import time
from pathlib import Path

import pymupdf

from backend.tasks import TaskCancelled

logger = logging.getLogger("vghtpe")
REPORT_TYPE_MAP = {
    "消防": "Fire_Equipment",
    "電力每日": "electricity_every_day",
    "電力每月": "electricity_every_month",
    "電力每周": "electricity_every_week",
    "排水每日": "drain_day",
    "排水每月": "drain_month",
    "排水每周": "drain_week",
}


def validate_month(value):
    try:
        parsed = datetime.datetime.strptime(value, "%Y-%m")
        return parsed.strftime("%Y-%m") == value and parsed.year >= 2022
    except (TypeError, ValueError):
        return False


def get_month_range(start_month, end_month):
    if not validate_month(start_month) or not validate_month(end_month):
        raise ValueError("月份格式須為 YYYY-MM，年份不得早於 2022")
    start = datetime.datetime.strptime(start_month, "%Y-%m")
    end = datetime.datetime.strptime(end_month, "%Y-%m")
    if start > end:
        raise ValueError("開始月份不得晚於結束月份")
    months = []
    current = start
    while True:
        months.append(current.strftime("%Y-%m"))
        if current == end:
            break
        current = current.replace(year=current.year + 1, month=1) if current.month == 12 else current.replace(month=current.month + 1)
    return months


def get_last_month(today=None):
    today = today or datetime.date.today()
    return (today.replace(day=1) - datetime.timedelta(days=1)).strftime("%Y-%m")


def build_report_url(report, report_type, month):
    if not validate_month(month):
        raise ValueError("無效的報表月份")
    if report_type == "消防":
        return f'https://vghtpe-ue.httc.com.tw/Report6{report["api_1"]}{month}{report["api_2"]}'
    if report_type not in REPORT_TYPE_MAP:
        raise ValueError("未知的報表類型")
    year, number = map(int, month.split("-"))
    last_day = calendar.monthrange(year, number)[1]
    return f'https://vghtpe-ue.httc.com.tw/Report6BatchAll{report["api_1"]}{month}-01/{month}-{last_day}{report["api_2"]}'


def catalog_reports(data):
    if not isinstance(data, dict):
        raise ValueError("data.json 最上層須為物件")
    reports = []
    for report_type, key in REPORT_TYPE_MAP.items():
        if not isinstance(data.get(key), list):
            raise ValueError(f"data.json 的 {key} 須為報表清單")
        for index, entry in enumerate(data[key]):
            if not isinstance(entry, dict) or not all(isinstance(entry.get(field), str) and entry[field].strip()
                                                     for field in ("name", "api_1", "api_2")):
                raise ValueError(f"data.json 的 {key} 第 {index + 1} 筆資料不完整")
            if any(ord(char) < 32 or char in '<>:"/\\|?*' for char in entry["name"]) or entry["name"].endswith((".", " ")):
                raise ValueError(f'報表名稱不可作為檔名：{entry["name"]}')
            reports.append({**entry, "id": f"{key}:{index}", "report_type": report_type, "category": report_type[:2]})
    return reports


def load_catalog(path):
    return catalog_reports(json.loads(Path(path).read_text(encoding="utf-8-sig")))


def valid_pdf(path):
    try:
        if not Path(path).is_file() or Path(path).stat().st_size == 0:
            return False
        with pymupdf.open(path) as document:
            return document.is_pdf and document.page_count > 0 and not document.needs_pass
    except (OSError, RuntimeError, ValueError):
        return False


class ReportDownloader:
    def __init__(self, executable, cancel_event=None, timeout=120):
        self.executable = str(executable)
        self.cancel_event = cancel_event or threading.Event()
        self.timeout = timeout

    def _check_cancelled(self):
        if self.cancel_event.is_set():
            raise TaskCancelled()

    @staticmethod
    def _terminate(process):
        if process.poll() is None:
            process.terminate()
        try:
            process.communicate(timeout=2)
        except subprocess.TimeoutExpired:
            process.kill()
            process.communicate()

    def _convert(self, url, destination):
        command = [self.executable, "--no-background", "--disable-javascript", "--encoding", "utf-8", "--quiet", url, str(destination)]
        process = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
                                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        deadline = time.monotonic() + self.timeout
        try:
            while True:
                self._check_cancelled()
                if time.monotonic() > deadline:
                    raise TimeoutError(f"報表轉換超過 {self.timeout} 秒")
                try:
                    _, stderr = process.communicate(timeout=0.2)
                    break
                except subprocess.TimeoutExpired:
                    continue
            if process.returncode != 0:
                detail = (stderr or b"").decode("utf-8", errors="replace").strip()[-1000:]
                raise RuntimeError(f"網頁轉 PDF 失敗：{detail or process.returncode}")
        finally:
            if process.poll() is None:
                self._terminate(process)

    def download_report(self, url, output_path, name, max_retries=3, log=None):
        log = log or (lambda message: None)
        self._check_cancelled()
        destination = Path(output_path)
        if valid_pdf(destination):
            log(f"{name} 已存在有效 PDF，跳過下載")
            return "skipped"
        destination.parent.mkdir(parents=True, exist_ok=True)
        for attempt in range(max_retries):
            self._check_cancelled()
            fd, temp = tempfile.mkstemp(prefix="report-", suffix=".pdf", dir=destination.parent)
            os.close(fd)
            try:
                log(f"正在下載 {name}（{attempt + 1}/{max_retries}）")
                self._convert(url, temp)
                self._check_cancelled()
                if not valid_pdf(temp):
                    raise RuntimeError("轉換結果不是有效的 PDF")
                os.replace(temp, destination)
                log(f"{name} 下載完成")
                return "downloaded"
            except TaskCancelled:
                raise
            except PermissionError as error:
                raise PermissionError("無法寫入 PDF，請關閉正在閱讀的檔案，並確認輸出資料夾可寫入") from error
            except (OSError, RuntimeError) as error:
                logger.exception("Download failed: %s", name)
                log(f"{name} 下載失敗：{error}")
                if attempt + 1 == max_retries:
                    return "failed"
                if self.cancel_event.wait((attempt + 1) * 5):
                    raise TaskCancelled()
            finally:
                Path(temp).unlink(missing_ok=True)
        return "failed"
