import unittest
from nippo_filename import year_month_from_filename

class TestFilenameParsing(unittest.TestCase):
    def test_mixed_filename_conventions(self):
        for name in ("2026_11__daily.xls", "【名店街】日報（2026年11月）.xls", "2026-11.xls"):
            with self.subTest(name=name):
                self.assertEqual(year_month_from_filename(name), (2026, 11))

    def test_invalid_name(self):
        with self.assertRaises(ValueError):
            year_month_from_filename("report.xls")
