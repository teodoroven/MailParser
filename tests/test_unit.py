import unittest

from main import (
    count_month,
    get_accurate,
    get_columns,
    get_date,
    get_max_date,
    get_row,
    parse_year,
    DATE_FORMAT_2,
)


class TestCoreUtils(unittest.TestCase):
    def test_get_accurate(self):
        string = "ООО «Рога и копыта, ко»"
        self.assertEqual(get_accurate(string, True), string)
        self.assertEqual(get_accurate(string, False), "ооо рога и копыта ко")

    def test_get_columns(self):
        self.assertEqual(get_columns(2), ["A", "B"])

    def test_get_max_date(self):
        self.assertEqual(
            get_max_date(["15.03.2025", "10.02.2025", "20.01.2025"]).strftime(DATE_FORMAT_2),
            "15.03.2025",
        )

    def test_count_month(self):
        self.assertEqual(count_month("01.12.2024"), 0)
        self.assertEqual(count_month("01.01.2025"), 1)
        self.assertEqual(count_month("01.03.2025"), 3)

    def test_parse_year(self):
        self.assertEqual(parse_year("01.12.", "01.01.2025"), "01.12.2024")
        self.assertEqual(parse_year("01.12.", "01.07.2025"), "01.12.2025")
        self.assertEqual(parse_year("01.12.", "01.01.2026"), "01.12.2025")

    def test_get_row(self):
        rows = ["", "ООО Рога и Копыта", "ООО «Свин и ко»"]
        self.assertEqual(get_row(rows, "ООО 'Рога и Копыта'", False), 1)
        self.assertEqual(get_row(rows, "ООО «Ромашка»", False), -1)

    def test_get_date(self):
        self.assertEqual(get_date(1), "31.01.2025")
        self.assertEqual(get_date(2), "28.02.2025")
        self.assertEqual(get_date(15), "31.03.2026")


if __name__ == "__main__":
    unittest.main()
