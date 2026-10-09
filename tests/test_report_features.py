import datetime
import json
import os
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

import pymupdf

from backend.config import get_data_file_path, load_settings, save_settings, open_folder
from backend.reports import ReportDownloader, build_report_url, get_last_month, get_month_range, validate_month, valid_pdf
from backend.tasks import TaskCancelled


def make_pdf(path, text="Report content"):
    with pymupdf.open() as document:
        page = document.new_page()
        page.insert_text((72, 72), text)
        document.save(path)


class DataFilePathTests(unittest.TestCase):
    def test_packaged_app_uses_data_json_next_to_executable(self):
        with mock.patch.object(sys, "frozen", True, create=True), mock.patch.object(sys, "executable", r"C:\Apps\default.exe"):
            self.assertEqual(r"C:\Apps\data.json", get_data_file_path())


class MonthRangeTests(unittest.TestCase):
    def test_two_year_inclusive_range(self):
        months = get_month_range("2022-01", "2023-12")
        self.assertEqual(24, len(months))
        self.assertEqual("2022-01", months[0])
        self.assertEqual("2023-12", months[-1])

    def test_single_month_range(self):
        self.assertEqual(["2024-02"], get_month_range("2024-02", "2024-02"))

    def test_invalid_ranges(self):
        for value in ("2021-12", "2022-00", "2022-13", "2022-1", "not-a-date"):
            self.assertFalse(validate_month(value))
        with self.assertRaises(ValueError):
            get_month_range("2023-01", "2022-12")

    def test_last_month_does_not_use_invalid_day(self):
        self.assertEqual("2026-02", get_last_month(datetime.date(2026, 3, 31)))
        self.assertEqual("2025-12", get_last_month(datetime.date(2026, 1, 31)))

    def test_last_supported_month_does_not_overflow(self):
        self.assertEqual(["9999-12"], get_month_range("9999-12", "9999-12"))


class ReportUrlTests(unittest.TestCase):
    def setUp(self):
        self.report = {"api_1": "/84/", "api_2": "/84/100"}

    def test_fire_report_url(self):
        self.assertEqual("https://vghtpe-ue.httc.com.tw/Report6/84/2023-12/84/100", build_report_url(self.report, "消防", "2023-12"))

    def test_batch_report_url_uses_leap_year_month_end(self):
        self.assertEqual("https://vghtpe-ue.httc.com.tw/Report6BatchAll/84/2024-02-01/2024-02-29/84/100", build_report_url(self.report, "電力每月", "2024-02"))

    def test_drain_report_uses_batch_endpoint(self):
        self.assertIn("2023-04-30", build_report_url(self.report, "排水每日", "2023-04"))

    def test_unknown_report_type_rejected(self):
        with self.assertRaises(ValueError):
            build_report_url(self.report, "電力未知", "2023-04")


class SettingsTests(unittest.TestCase):
    def test_missing_and_corrupt_settings_default_to_disabled(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            self.assertFalse(load_settings(path)["open_folder_after_completion"])
            path.write_text("not json", encoding="utf-8")
            self.assertFalse(load_settings(path)["open_folder_after_completion"])
            path.write_text("[]", encoding="utf-8")
            self.assertFalse(load_settings(path)["open_folder_after_completion"])

    def test_setting_is_saved_and_loaded(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            settings = load_settings(path)
            settings["open_folder_after_completion"] = True
            settings["output_folder"] = str(Path(directory) / "reports")
            save_settings(settings, path)
            self.assertEqual(settings, load_settings(path))

    def test_failed_write_preserves_previous_settings(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            save_settings({"open_folder_after_completion": False}, path)
            before = path.read_bytes()
            with mock.patch("backend.config.os.replace", side_effect=PermissionError):
                with self.assertRaises(PermissionError):
                    save_settings({"open_folder_after_completion": True}, path)
            self.assertEqual(before, path.read_bytes())
            self.assertEqual([path], list(Path(directory).iterdir()))

    def test_explicit_open_folder(self):
        with tempfile.TemporaryDirectory() as directory, mock.patch("backend.config.os.startfile") as startfile:
            folder = Path(directory) / "reports"
            open_folder(folder)
            startfile.assert_called_once_with(str(folder.resolve()))
            self.assertTrue(folder.is_dir())


class DownloadTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.output = Path(self.temp.name) / "report.pdf"
        self.downloader = ReportDownloader("mock-executable")

    def test_valid_existing_pdf_is_skipped(self):
        make_pdf(self.output)
        with mock.patch.object(self.downloader, "_convert") as convert:
            self.assertEqual("skipped", self.downloader.download_report("mock", self.output, "report"))
            convert.assert_not_called()

    def test_non_pdf_residual_is_redownloaded(self):
        self.output.write_bytes(b"incomplete non-PDF")
        with mock.patch.object(self.downloader, "_convert", side_effect=lambda url, path: make_pdf(path)) as convert:
            self.assertEqual("downloaded", self.downloader.download_report("mock", self.output, "report"))
            convert.assert_called_once()
        self.assertTrue(valid_pdf(self.output))

    def test_failed_conversion_does_not_leave_formal_output(self):
        def fail(url, path):
            Path(path).write_bytes(b"partial output")
            raise RuntimeError("conversion failed")
        with mock.patch.object(self.downloader, "_convert", side_effect=fail):
            self.assertEqual("failed", self.downloader.download_report("mock", self.output, "report", max_retries=1))
        self.assertFalse(self.output.exists())
        self.assertEqual([], list(Path(self.temp.name).iterdir()))

    def test_empty_existing_file_is_redownloaded(self):
        self.output.touch()
        with mock.patch.object(self.downloader, "_convert", side_effect=lambda url, path: make_pdf(path)):
            self.assertEqual("downloaded", self.downloader.download_report("mock", self.output, "report"))

    def test_locked_destination_has_friendly_error_and_cleans_temp(self):
        with mock.patch.object(self.downloader, "_convert", side_effect=lambda url, path: make_pdf(path)), mock.patch("backend.reports.os.replace", side_effect=PermissionError):
            with self.assertRaisesRegex(PermissionError, "請關閉"):
                self.downloader.download_report("mock", self.output, "report")
        self.assertFalse(self.output.exists())
        self.assertEqual([], list(Path(self.temp.name).iterdir()))

    def test_cancelled_task_does_not_skip_existing_pdf(self):
        make_pdf(self.output)
        self.downloader.cancel_event.set()
        with self.assertRaises(TaskCancelled):
            self.downloader.download_report("mock", self.output, "report")

    def test_cancel_during_retry_is_immediate(self):
        self.downloader.cancel_event.wait = mock.Mock(return_value=True)
        with mock.patch.object(self.downloader, "_convert", side_effect=RuntimeError("failed")):
            with self.assertRaises(TaskCancelled):
                self.downloader.download_report("mock", self.output, "report")

    def test_kill_after_terminate_timeout_is_reaped(self):
        process = mock.Mock()
        process.poll.return_value = None
        process.communicate.side_effect = [__import__("subprocess").TimeoutExpired("mock", 2), (None, b"")]
        ReportDownloader._terminate(process)
        process.terminate.assert_called_once()
        process.kill.assert_called_once()
        self.assertEqual(2, process.communicate.call_count)


class ConverterIntegrationTests(unittest.TestCase):
    @unittest.skipUnless(sys.platform == "win32", "Windows converter")
    def test_bundled_converter_produces_searchable_pdf(self):
        executable = Path(__file__).resolve().parent.parent / "wkhtmltopdf.exe"
        with tempfile.TemporaryDirectory() as directory:
            html = Path(directory) / "input.html"
            html.write_text("<html><body><h1>Converter smoke test</h1></body></html>", encoding="utf-8")
            output = Path(directory) / "output.pdf"
            result = ReportDownloader(executable).download_report(html.as_uri(), output, "smoke", max_retries=1)
            self.assertEqual("downloaded", result)
            with pymupdf.open(output) as document:
                self.assertIn("Converter smoke test", document[0].get_text())

    def test_conversion_timeout_terminates_and_reaps_child(self):
        import subprocess
        process = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"], stdout=subprocess.DEVNULL,
                                   stderr=subprocess.PIPE, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        downloader = ReportDownloader("mock", timeout=0.01)
        try:
            with mock.patch("backend.reports.subprocess.Popen", return_value=process):
                with self.assertRaises(TimeoutError):
                    downloader._convert("mock", "mock")
            self.assertIsNotNone(process.poll())
        finally:
            if process.poll() is None:
                process.kill()
                process.communicate()


if __name__ == "__main__":
    unittest.main()
