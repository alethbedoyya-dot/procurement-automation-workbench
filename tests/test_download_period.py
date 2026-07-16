import unittest
from datetime import date

from download_period import monthly_download_dir, previous_month_data_folder_name


class DownloadPeriodTests(unittest.TestCase):
    def test_uses_the_previous_natural_month_with_its_year(self):
        self.assertEqual(
            previous_month_data_folder_name(date(2026, 7, 16)),
            "2026年六月数据",
        )
        self.assertEqual(
            previous_month_data_folder_name(date(2026, 8, 1)),
            "2026年七月数据",
        )

    def test_crosses_the_year_boundary_in_january(self):
        self.assertEqual(
            previous_month_data_folder_name(date(2027, 1, 1)),
            "2026年十二月数据",
        )

    def test_places_the_period_between_downloads_and_category(self):
        root = monthly_download_dir(r"C:\work\downloads", date(2026, 7, 16))

        self.assertEqual(root, r"C:\work\downloads\2026年六月数据")


if __name__ == "__main__":
    unittest.main()
