import copy
import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

from fastapi.testclient import TestClient

from backend.app import create_app
from backend.catalog import CatalogConflict, CatalogStore
from backend.reports import REPORT_TYPE_MAP


def sample_data():
    data = {key: [] for key in REPORT_TYPE_MAP.values()}
    data.update({"squadName": "測試巡檢", "Building_name": ["第一大樓"], "custom": {"keep": True}})
    data["Fire_Equipment"] = [
        {"name": "第一份報表", "api_1": "/16/", "api_2": "/16/32", "extra": "保留"},
        {"name": "第二份報表", "api_1": "/17/", "api_2": "/17/33"},
    ]
    return data


def report(name="新增報表", report_type="消防"):
    return {"name": name, "report_type": report_type, "api_1": "/20/", "api_2": "/20/40"}


class CatalogTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "data.json"
        self.original = sample_data()
        self.path.write_text(json.dumps(self.original, ensure_ascii=False), encoding="utf-8")
        self.store = CatalogStore(self.path)
        self.revision = self.store.snapshot()["revision"]

    def read(self):
        return json.loads(self.path.read_text(encoding="utf-8"))

    def test_add_preserves_metadata_and_exact_previous_backup(self):
        before = self.path.read_bytes()
        snapshot = self.store.mutate(self.revision, report())
        self.assertEqual(3, len(snapshot["reports"]))
        self.assertNotEqual(self.revision, snapshot["revision"])
        self.assertEqual(before, self.store.backup_path.read_bytes())
        for key in ("Building_name", "squadName", "custom"):
            self.assertEqual(self.original[key], self.read()[key])
        self.assertEqual(snapshot, CatalogStore(self.path).snapshot())

    def test_edit_move_and_delete_preserve_entry_extras(self):
        updated = self.store.mutate(self.revision, report("更新名稱", "電力每日"), "Fire_Equipment:0")
        self.assertEqual("保留", self.read()["electricity_every_day"][0]["extra"])
        self.assertEqual("第二份報表", self.read()["Fire_Equipment"][0]["name"])
        result = self.store.mutate(updated["revision"], report_id="electricity_every_day:0", delete=True)
        self.assertEqual(1, len(result["reports"]))
        self.assertEqual([], self.read()["electricity_every_day"])

    def test_edit_within_type_and_backup_rotates(self):
        first = self.store.mutate(self.revision, report("更新名稱"), "Fire_Equipment:0")
        before = self.path.read_bytes()
        self.assertEqual("保留", self.read()["Fire_Equipment"][0]["extra"])
        self.store.mutate(first["revision"], report("再新增"))
        self.assertEqual(before, self.store.backup_path.read_bytes())

    def test_duplicate_name_in_same_category_rejected_but_other_category_allowed(self):
        with self.assertRaisesRegex(ValueError, "相同報表名稱"):
            self.store.mutate(self.revision, report("第一份報表"))
        self.store.mutate(self.revision, report("第一份報表", "電力每日"))

    def test_duplicate_name_across_electricity_frequencies_rejected(self):
        first = self.store.mutate(self.revision, report("電力報表", "電力每日"))
        with self.assertRaisesRegex(ValueError, "相同報表名稱"):
            self.store.mutate(first["revision"], report("電力報表", "電力每月"))

    def test_invalid_names_and_parameters_leave_original_untouched(self):
        before = self.path.read_bytes()
        invalid = [
            {"name": " "}, {"name": "CON"}, {"name": "a/b"}, {"name": "後綴."}, {"name": "\x01name"},
            {"name": " 前綴"}, {"api_1": "https://example.org"}, {"api_2": "/16/abc"}, {"report_type": "新增類型"},
        ]
        for change in invalid:
            with self.subTest(change=change), self.assertRaises(ValueError):
                self.store.mutate(self.revision, {**report(), **change})
        self.assertEqual(before, self.path.read_bytes())
        self.assertFalse(self.store.backup_path.exists())

    def test_legacy_parameter_format_remains_supported(self):
        self.store.mutate(self.revision, {**report(), "api_2": "417/453"})
        self.assertEqual("417/453", self.read()["Fire_Equipment"][-1]["api_2"])

    def test_stale_revision_cannot_delete_shifted_index(self):
        updated = self.store.mutate(self.revision, report_id="Fire_Equipment:0", delete=True)
        with self.assertRaises(CatalogConflict):
            self.store.mutate(self.revision, report_id="Fire_Equipment:0", delete=True)
        self.assertEqual("第二份報表", updated["reports"][0]["name"])
        self.assertEqual(1, len(self.store.snapshot()["reports"]))

    def test_external_edits_are_reloaded_and_reject_stale_write(self):
        changed = copy.deepcopy(self.original)
        changed["custom"]["external"] = "keep"
        self.path.write_text(json.dumps(changed), encoding="utf-8")
        self.assertNotEqual(self.revision, self.store.snapshot()["revision"])
        with self.assertRaises(CatalogConflict):
            self.store.mutate(self.revision, report())
        self.assertEqual(changed, self.read())

    def test_backup_failure_leaves_data_unchanged(self):
        before = self.path.read_bytes()
        with mock.patch("backend.catalog.os.replace", side_effect=PermissionError):
            with self.assertRaises(PermissionError):
                self.store.mutate(self.revision, report())
        self.assertEqual(before, self.path.read_bytes())
        self.assertEqual([self.path], list(self.path.parent.iterdir()))

    def test_data_replace_failure_retains_original_and_backup(self):
        before = self.path.read_bytes()
        import os
        replace = os.replace
        def fail_data(source, destination):
            if Path(destination) == self.path:
                raise PermissionError("locked")
            replace(source, destination)
        with mock.patch("backend.catalog.os.replace", side_effect=fail_data):
            with self.assertRaises(PermissionError):
                self.store.mutate(self.revision, report())
        self.assertEqual(before, self.path.read_bytes())
        self.assertEqual(before, self.store.backup_path.read_bytes())
        self.assertEqual(self.revision, self.store.snapshot()["revision"])
        self.assertFalse(list(self.path.parent.glob("*.tmp")))

    def test_missing_report_and_corrupt_structure_are_rejected(self):
        with self.assertRaises(KeyError):
            self.store.mutate(self.revision, report(), "Fire_Equipment:99")
        for data in ([], {"Fire_Equipment": []}, {**self.original, "Fire_Equipment": [None]}):
            self.path.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(ValueError):
                self.store.snapshot()


class CatalogApiTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.path = self.directory / "data.json"
        self.path.write_text(json.dumps(sample_data(), ensure_ascii=False), encoding="utf-8")
        self.app = create_app(data_path=self.path, settings_path=self.directory / "settings.json")
        self.client = TestClient(self.app)
        self.client.__enter__()
        self.addCleanup(self.client.__exit__, None, None, None)
        self.client.headers["X-App-Token"] = self.client.get("/api/bootstrap").json()["token"]

    def catalog(self):
        response = self.client.get("/api/catalog")
        self.assertEqual(200, response.status_code)
        return response.json()

    def test_crud_synchronizes_bootstrap_and_survives_restart(self):
        initial = self.catalog()
        added = self.client.post("/api/catalog/reports", json={**report(), "revision": initial["revision"]})
        self.assertEqual(201, added.status_code)
        catalog = added.json()
        item = next(row for row in catalog["reports"] if row["name"] == "新增報表")
        changed = self.client.put("/api/catalog/reports/" + item["id"], json={**report("已修改", "排水每月"), "revision": catalog["revision"]})
        self.assertEqual(200, changed.status_code)
        catalog = changed.json()
        item = next(row for row in catalog["reports"] if row["name"] == "已修改")
        bootstrap = self.client.get("/api/bootstrap").json()
        self.assertEqual(catalog["revision"], bootstrap["catalog_revision"])
        self.assertEqual(3, len(bootstrap["reports"]))
        recreated = create_app(data_path=self.path, settings_path=self.directory / "settings.json")
        self.assertEqual(catalog["revision"], recreated.state.runtime.catalog.snapshot()["revision"])
        response = self.client.request("DELETE", "/api/catalog/reports/" + item["id"], json={"revision": catalog["revision"]})
        self.assertEqual(200, response.status_code)
        self.assertEqual(2, len(response.json()["reports"]))

    def test_session_and_validation_and_missing_id(self):
        initial = self.catalog()
        self.assertEqual(403, self.client.get("/api/catalog", headers={"X-App-Token": "wrong"}).status_code)
        for input in ({**report(), "revision": initial["revision"], "api_1": "/bad/"}, {**report(), "revision": ""}):
            self.assertEqual(422, self.client.post("/api/catalog/reports", json=input).status_code)
        self.assertEqual(404, self.client.request("DELETE", "/api/catalog/reports/missing", json={"revision": initial["revision"]}).status_code)

    def test_stale_mutation_and_stale_download_are_rejected(self):
        initial = self.catalog()
        self.client.post("/api/catalog/reports", json={**report(), "revision": initial["revision"]})
        self.assertEqual(409, self.client.request("DELETE", "/api/catalog/reports/Fire_Equipment:0", json={"revision": initial["revision"]}).status_code)
        self.assertEqual(409, self.client.post("/api/tasks/download", json={
            "report_ids": ["Fire_Equipment:0"], "start_month": "2026-01", "end_month": "2026-01",
            "catalog_revision": initial["revision"],
        }).status_code)

    def test_read_and_write_failures_return_friendly_errors(self):
        initial = self.catalog()
        with mock.patch("backend.catalog.os.replace", side_effect=PermissionError):
            response = self.client.post("/api/catalog/reports", json={**report(), "revision": initial["revision"]})
        self.assertEqual(400, response.status_code)
        self.assertIn("原資料未替換", response.json()["detail"])
        self.path.write_text("invalid", encoding="utf-8")
        self.assertEqual(400, self.client.get("/api/catalog").status_code)
        self.assertEqual(400, self.client.get("/api/bootstrap").status_code)

    def test_running_download_uses_original_snapshot_while_catalog_changes(self):
        initial = self.catalog()
        started = threading.Event()
        release = threading.Event()
        names = []
        def download(downloader, url, output, name, max_retries=3, log=None):
            names.append(name)
            started.set()
            release.wait(3)
            return "skipped"
        with mock.patch("backend.app.ReportDownloader.download_report", autospec=True, side_effect=download):
            try:
                response = self.client.post("/api/tasks/download", json={
                    "report_ids": ["Fire_Equipment:0"], "start_month": "2026-01", "end_month": "2026-02",
                    "catalog_revision": initial["revision"],
                })
                self.assertEqual(202, response.status_code)
                self.assertTrue(started.wait(2))
                edited = self.client.put("/api/catalog/reports/Fire_Equipment:0", json={**report("任務期間修改"), "revision": initial["revision"]})
                self.assertEqual(200, edited.status_code)
            finally:
                release.set()
                self.app.state.runtime.tasks.worker.join(3)
        self.assertEqual(["第一份報表", "第一份報表"], names)
        self.assertEqual("completed", self.app.state.runtime.tasks.snapshot()["status"])
