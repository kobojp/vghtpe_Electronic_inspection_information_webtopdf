import tempfile
import threading
import time
import unittest
from pathlib import Path

import pymupdf

from backend.pdf_core import extract_pdf
from backend.tasks import TaskBusy, TaskCancelled, TaskManager, TaskContext
from tests.test_report_features import make_pdf


class TaskManagerTests(unittest.TestCase):
    def test_cancelling_keeps_task_busy_until_worker_exits(self):
        manager = TaskManager()
        entered = threading.Event()
        release = threading.Event()
        events = []
        def work(context):
            events.append(context.cancel_event)
            entered.set()
            release.wait(2)
            context.check_cancelled()
        first = manager.start("download", "test", 3, ".", work)
        self.assertTrue(entered.wait(1))
        manager.cancel(first["id"])
        self.assertEqual("cancelling", manager.snapshot()["status"])
        with self.assertRaises(TaskBusy):
            manager.start("download", "second", 1, ".", work)
        self.assertTrue(events[0].is_set())
        release.set()
        manager.worker.join(2)
        self.assertEqual("cancelled", manager.snapshot()["status"])
        self.assertEqual(3, manager.snapshot()["cancelled"])
        next_events = []
        manager.start("download", "next", 1, ".", lambda context: next_events.append(context.cancel_event))
        manager.worker.join(2)
        self.assertIsNot(events[0], next_events[0])
        self.assertFalse(next_events[0].is_set())

    def test_counts_distinguish_skipped_and_failed(self):
        manager = TaskManager()
        def work(context):
            for result in ("downloaded", "skipped", "failed"):
                context.record(result)
        manager.start("download", "test", 3, ".", work)
        manager.worker.join(2)
        task = manager.snapshot()
        self.assertEqual((1, 1, 1), (task["successful"], task["skipped"], task["failed"]))
        self.assertEqual(100, task["progress"])
        task["logs"].clear()
        self.assertTrue(manager.snapshot()["logs"])

    def test_shutdown_cancels_active_work(self):
        manager = TaskManager()
        entered = threading.Event()
        def work(context):
            entered.set()
            context.cancel_event.wait(2)
            context.check_cancelled()
        manager.start("extract", "test", 1, ".", work)
        self.assertTrue(entered.wait(1))
        manager.shutdown()
        self.assertFalse(manager.worker.is_alive())
        self.assertEqual("cancelled", manager.snapshot()["status"])


class PdfExtractionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.manager = TaskManager()

    def extract(self, paths, keyword="TARGET", keep=False):
        output = self.directory / "output"
        results = []
        def work(context):
            for path in paths:
                results.append(extract_pdf(path, output, keyword, keep, context))
                context.record(results[-1])
        self.manager.start("extract", "test", len(paths), str(output), work)
        self.manager.worker.join(5)
        self.assertEqual("completed", self.manager.snapshot()["status"])
        return output, results

    def test_same_filename_across_months_preserves_both_outputs(self):
        sources = []
        for month, content in (("2026-01", "TARGET January"), ("2026-02", "TARGET February")):
            folder = self.directory / month
            folder.mkdir()
            source = folder / "report.pdf"
            make_pdf(source, content)
            sources.append(source)
        output, results = self.extract(sources)
        files = sorted(output.glob("*.pdf"))
        self.assertEqual(2, len(files))
        texts = []
        for path in files:
            with pymupdf.open(path) as document:
                texts.append(document[0].get_text())
        self.assertTrue(any("January" in text for text in texts))
        self.assertTrue(any("February" in text for text in texts))

    def test_existing_output_is_not_overwritten(self):
        source = self.directory / "report.pdf"
        make_pdf(source, "TARGET first")
        output, _ = self.extract([source])
        before = (output / "report.pdf").read_bytes()
        make_pdf(source, "TARGET second")
        self.extract([source])
        self.assertEqual(before, (output / "report.pdf").read_bytes())
        self.assertTrue((output / "report__2.pdf").exists())

    def test_whitespace_and_case_match(self):
        source = self.directory / "report.pdf"
        make_pdf(source, "T A R G E T")
        output, results = self.extract([source], "target")
        self.assertEqual(["extracted"], results)
        self.assertEqual(1, len(list(output.glob("*.pdf"))))

    def test_no_match_does_not_output_contractor_page_alone(self):
        source = self.directory / "report.pdf"
        with pymupdf.open() as document:
            page = document.new_page()
            page.insert_text((72, 72), "承商駐院工程師", fontname="china-t")
            document.save(source)
        output, results = self.extract([source], keep=True)
        self.assertEqual(["skipped"], results)
        self.assertFalse(output.exists())

    def test_contractor_page_preserves_order_without_duplicate_pages(self):
        source = self.directory / "report.pdf"
        with pymupdf.open() as document:
            document.new_page().insert_text((72, 72), "承商駐院工程師", fontname="china-t")
            document.new_page().insert_text((72, 72), "TARGET")
            document.save(source)
        output, _ = self.extract([source], keep=True)
        with pymupdf.open(output / "report.pdf") as document:
            self.assertEqual(2, document.page_count)
            self.assertIn("承商駐院工程師", document[0].get_text())
            self.assertIn("TARGET", document[1].get_text())

    def test_cancelled_extraction_has_no_output(self):
        source = self.directory / "report.pdf"
        make_pdf(source, "TARGET")
        output = self.directory / "output"
        def work(context):
            context.cancel_event.set()
            extract_pdf(source, output, "TARGET", False, context)
        self.manager.start("extract", "test", 1, str(output), work)
        self.manager.worker.join(2)
        self.assertEqual("cancelled", self.manager.snapshot()["status"])
        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
