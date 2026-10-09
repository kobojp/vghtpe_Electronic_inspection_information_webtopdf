import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

from fastapi.testclient import TestClient

from backend.app import create_app
from backend.config import save_settings
from tests.test_report_features import make_pdf


class ApiTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp.name)
        self.settings_path = self.directory / "settings.json"
        save_settings({"output_folder": str(self.directory / "reports"), "extract_folder": str(self.directory / "extracts"), "open_folder_after_completion": False}, self.settings_path)
        self.dialogs = mock.Mock()
        self.dialogs.select_pdfs.return_value = []
        self.dialogs.select_folder.return_value = None
        self.app = create_app(settings_path=self.settings_path, dialogs=self.dialogs)
        self.client = TestClient(self.app)
        self.client.__enter__()
        self.bootstrap = self.client.get("/api/bootstrap").json()
        self.client.headers["X-App-Token"] = self.bootstrap["token"]

    def tearDown(self):
        self.client.__exit__(None, None, None)
        self.temp.cleanup()

    def test_catalog_has_all_existing_report_types(self):
        self.assertEqual(378, len(self.bootstrap["reports"]))
        self.assertEqual(7, len(self.bootstrap["report_types"]))
        self.assertTrue(self.bootstrap["native_dialogs"])

    def test_mutations_require_session_token(self):
        response = self.client.post("/api/desktop/select-folder", headers={"X-App-Token": "wrong"})
        self.assertEqual(403, response.status_code)
        self.dialogs.select_folder.assert_not_called()

    def test_foreign_origin_cannot_read_bootstrap(self):
        response = self.client.get("/api/bootstrap", headers={"Origin": "https://example.invalid"})
        self.assertEqual(403, response.status_code)

    def test_invalid_dates_and_unknown_ids_rejected(self):
        item = self.bootstrap["reports"][0]
        for body in ({"report_ids": [item["id"]], "start_month": "2026-03", "end_month": "2026-02"},
                     {"report_ids": ["missing"], "start_month": "2026-02", "end_month": "2026-02"}):
            self.assertEqual(422, self.client.post("/api/tasks/download", json=body).status_code)
        self.assertIsNone(self.app.state.runtime.tasks.snapshot())

    def test_download_uses_month_folders_and_separate_skip_count(self):
        item = self.bootstrap["reports"][0]
        paths = []
        def convert(downloader, url, path, name, max_retries=3, log=None):
            paths.append(Path(path))
            return "downloaded" if len(paths) == 1 else "skipped"
        with mock.patch("backend.app.ReportDownloader.download_report", autospec=True, side_effect=convert):
            response = self.client.post("/api/tasks/download", json={"report_ids": [item["id"]], "start_month": "2026-01", "end_month": "2026-02"})
            self.assertEqual(202, response.status_code)
            self.app.state.runtime.tasks.worker.join(2)
        task = self.client.get("/api/tasks/current").json()
        self.assertEqual((1, 1, 0), (task["successful"], task["skipped"], task["failed"]))
        self.assertEqual(["2026-01", "2026-02"], [path.parent.name for path in paths])
        self.assertEqual("completed", task["status"])

    def test_native_dialog_cancellation_is_empty_result(self):
        self.assertEqual({"paths": []}, self.client.post("/api/desktop/select-pdfs").json())
        self.assertEqual({"path": None}, self.client.post("/api/desktop/select-folder").json())

    def test_settings_persist_and_invalid_path_is_rejected(self):
        settings = self.bootstrap["settings"]
        settings["open_folder_after_completion"] = True
        self.assertEqual(settings, self.client.put("/api/settings", json=settings).json())
        recreated = create_app(settings_path=self.settings_path)
        self.assertTrue(recreated.state.runtime.settings["open_folder_after_completion"])
        settings["output_folder"] = "relative"
        self.assertEqual(422, self.client.put("/api/settings", json=settings).status_code)

    def test_auto_open_folder_obeys_preference(self):
        report = self.bootstrap["reports"][0]
        body = {"report_ids": [report["id"]], "start_month": "2026-01", "end_month": "2026-01"}
        with mock.patch("backend.app.ReportDownloader.download_report", return_value="skipped"), mock.patch("backend.app.config.open_folder") as open_folder:
            self.client.post("/api/tasks/download", json=body)
            self.app.state.runtime.tasks.worker.join(2)
            open_folder.assert_not_called()
            settings = self.bootstrap["settings"]
            settings["open_folder_after_completion"] = True
            self.client.put("/api/settings", json=settings)
            self.client.post("/api/tasks/download", json=body)
            self.app.state.runtime.tasks.worker.join(2)
            open_folder.assert_called_once()

    def test_pdf_api_extracts_real_document(self):
        source = self.directory / "source.pdf"
        make_pdf(source, "TARGET")
        response = self.client.post("/api/tasks/extract", json={"paths": [str(source)], "keyword": "TARGET", "keep_contractor_page": False})
        self.assertEqual(202, response.status_code)
        self.app.state.runtime.tasks.worker.join(3)
        task = self.client.get("/api/tasks/current").json()
        self.assertEqual(1, task["successful"])
        self.assertTrue(list(Path(task["output_folder"]).glob("*.pdf")))

    def test_missing_or_empty_pdf_input_rejected(self):
        self.assertEqual(422, self.client.post("/api/tasks/extract", json={"paths": [], "keyword": "TARGET"}).status_code)
        self.assertEqual(422, self.client.post("/api/tasks/extract", json={"paths": [str(self.directory / "missing.pdf")], "keyword": "TARGET"}).status_code)

    def test_busy_task_rejected_until_cancellation_finishes(self):
        entered, release = threading.Event(), threading.Event()
        def work(context):
            entered.set()
            release.wait(2)
            context.check_cancelled()
        task = self.app.state.runtime.tasks.start("download", "test", 1, str(self.directory), work)
        self.assertTrue(entered.wait(1))
        try:
            self.assertEqual("cancelling", self.client.post(f'/api/tasks/{task["id"]}/cancel').json()["status"])
            response = self.client.post("/api/tasks/download", json={"report_ids": [self.bootstrap["reports"][0]["id"]], "start_month": "2026-01", "end_month": "2026-01"})
            self.assertEqual(409, response.status_code)
        finally:
            release.set()
            self.app.state.runtime.tasks.worker.join(2)
        self.assertEqual("cancelled", self.client.get("/api/tasks/current").json()["status"])


if __name__ == "__main__":
    unittest.main()
