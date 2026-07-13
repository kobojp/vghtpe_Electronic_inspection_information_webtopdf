import json
import os
import tempfile
import unittest
from unittest import mock

from main import (
    build_report_url,
    get_month_range,
    get_smooth_progress_value,
    htmltopdf,
    load_settings,
    save_settings,
    validate_month,
)


class MonthRangeTests(unittest.TestCase):
    def test_two_year_inclusive_range(self):
        months = get_month_range("2022-01", "2023-12")

        self.assertEqual(24, len(months))
        self.assertEqual("2022-01", months[0])
        self.assertEqual("2023-12", months[-1])

    def test_single_month_range(self):
        self.assertEqual(["2024-02"], get_month_range("2024-02", "2024-02"))

    def test_invalid_ranges(self):
        invalid_months = ["2021-12", "2022-00", "2022-13", "2022-1", "not-a-date"]
        for value in invalid_months:
            with self.subTest(value=value):
                self.assertFalse(validate_month(value))

        with self.assertRaises(ValueError):
            get_month_range("2023-01", "2022-12")


class SmoothProgressTests(unittest.TestCase):
    def test_progress_moves_toward_target_without_large_jump(self):
        first_value = get_smooth_progress_value(0, 100)

        self.assertGreater(first_value, 0)
        self.assertLessEqual(first_value, 2.5)

    def test_progress_never_overshoots_target(self):
        self.assertEqual(10, get_smooth_progress_value(9.9, 10))
        self.assertEqual(25, get_smooth_progress_value(25, 25))


class ReportUrlTests(unittest.TestCase):
    def setUp(self):
        self.report = {"api_1": "/84/", "api_2": "/84/100"}

    def test_fire_report_url(self):
        self.assertEqual(
            "https://vghtpe-ue.httc.com.tw/Report6/84/2023-12/84/100",
            build_report_url(self.report, "消防", "2023-12"),
        )

    def test_batch_report_url_uses_leap_year_month_end(self):
        self.assertEqual(
            "https://vghtpe-ue.httc.com.tw/Report6BatchAll/84/2024-02-01/2024-02-29/84/100",
            build_report_url(self.report, "電力每月", "2024-02"),
        )

    def test_drain_report_uses_batch_endpoint(self):
        url = build_report_url(self.report, "排水每日", "2023-04")

        self.assertIn("Report6BatchAll", url)
        self.assertIn("2023-04-30", url)


class SettingsTests(unittest.TestCase):
    def test_missing_and_corrupt_settings_default_to_disabled(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, "settings.json")
            self.assertFalse(load_settings(path)["open_folder_after_completion"])

            with open(path, "w", encoding="utf-8") as settings_file:
                settings_file.write("not json")
            self.assertFalse(load_settings(path)["open_folder_after_completion"])

    def test_setting_is_saved_and_loaded(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, "settings.json")
            save_settings(True, path)

            self.assertTrue(load_settings(path)["open_folder_after_completion"])
            with open(path, encoding="utf-8") as settings_file:
                self.assertEqual(
                    {"open_folder_after_completion": True},
                    json.load(settings_file),
                )


class OpenFolderTests(unittest.TestCase):
    def test_startfile_obeys_open_folder_setting(self):
        handler = object.__new__(htmltopdf)
        handler.open_folder_after_completion = False

        with mock.patch("main.os.startfile") as startfile:
            handler.startfile("output-folder")
            startfile.assert_not_called()

            handler.open_folder_after_completion = True
            handler.startfile("output-folder")
            startfile.assert_called_once_with("output-folder")

    def test_explicit_open_folder_ignores_automatic_setting(self):
        handler = object.__new__(htmltopdf)
        handler.open_folder_after_completion = False

        with mock.patch("main.os.startfile") as startfile:
            handler.open_folder("output-folder")

        startfile.assert_called_once_with("output-folder")


class ExistingReportTests(unittest.TestCase):
    def test_non_empty_existing_pdf_is_skipped(self):
        handler = object.__new__(htmltopdf)
        messages = []
        handler.progress_callback = messages.append

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = os.path.join(temp_dir, "二門診滅火器(月)檢查.pdf")
            with open(output_path, "wb") as output_file:
                output_file.write(b"existing pdf")

            with mock.patch("main.subprocess.Popen") as popen:
                result = handler.download_report(
                    "https://example.invalid/report",
                    output_path,
                    "二門診滅火器(月)檢查",
                )

            self.assertTrue(result)
            popen.assert_not_called()
            self.assertTrue(any("已存在，跳過下載" in message for message in messages))


if __name__ == "__main__":
    unittest.main()

