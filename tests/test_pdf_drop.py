import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from desktop.drop import bind_pdf_drop, dropped_pdf_files
from tests.test_report_features import make_pdf


class PdfDropTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.first = self.directory / "第一份.PDF"
        self.second = self.directory / "第二份.pdf"
        make_pdf(self.first, "TARGET first")
        make_pdf(self.second, "TARGET second")

    def item(self, path):
        return {"name": path.name, "pywebviewFullPath": str(path)}

    def test_native_drop_accepts_multiple_pdfs_and_deduplicates(self):
        event = {"dataTransfer": {"files": [self.item(self.first), self.item(self.second), self.item(self.first)]}}
        self.assertEqual({"paths": [str(self.first.resolve()), str(self.second.resolve())], "rejected": []}, dropped_pdf_files(event))

    def test_non_pdf_missing_corrupt_directory_and_unresolved_paths_rejected(self):
        text = self.directory / "note.txt"
        text.write_text("text", encoding="utf-8")
        corrupt = self.directory / "broken.pdf"
        corrupt.write_bytes(b"not a PDF")
        folder = self.directory / "folder.pdf"
        folder.mkdir()
        missing = self.directory / "missing.pdf"
        result = dropped_pdf_files({"dataTransfer": {"files": [
            self.item(self.first), self.item(text), self.item(corrupt), self.item(folder), self.item(missing),
            {"name": "no-path.pdf"}, {"name": "relative.pdf", "pywebviewFullPath": "relative.pdf"},
        ]}})
        self.assertEqual([str(self.first.resolve())], result["paths"])
        self.assertEqual(["note.txt", "broken.pdf", "folder.pdf", "missing.pdf", "no-path.pdf", "relative.pdf"], result["rejected"])

    def test_empty_text_drop_is_ignored(self):
        self.assertEqual({"paths": [], "rejected": []}, dropped_pdf_files({"dataTransfer": {"files": []}}))

    def test_bind_dispatches_complete_paths_as_safe_json(self):
        window = mock.Mock()
        zone = window.dom.get_element.return_value
        window.evaluate_js.side_effect = [False, None, True, None]
        bind_pdf_drop(window)
        zone.on.assert_called_once()
        handler = zone.on.call_args.args[1]
        self.assertTrue(handler.prevent_default)
        handler.callback({"dataTransfer": {"files": [self.item(self.first)]}})
        script = window.evaluate_js.call_args.args[0]
        self.assertIn("vghtpe:pdf-drop", script)
        payload = script.split("detail: ", 1)[1].rsplit(" }", 1)[0]
        self.assertEqual([str(self.first.resolve())], json.loads(payload)["paths"])

    def test_busy_or_hidden_zone_does_not_dispatch(self):
        window = mock.Mock()
        window.evaluate_js.side_effect = [False, None, False]
        bind_pdf_drop(window)
        handler = window.dom.get_element.return_value.on.call_args.args[1]
        handler.callback({"dataTransfer": {"files": [self.item(self.first)]}})
        self.assertEqual(3, window.evaluate_js.call_count)

    def test_already_bound_zone_is_not_registered_again(self):
        window = mock.Mock()
        window.evaluate_js.return_value = True
        bind_pdf_drop(window)
        window.dom.get_element.return_value.on.assert_not_called()
