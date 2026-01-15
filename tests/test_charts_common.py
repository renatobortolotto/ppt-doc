import unittest

from utils.charts_common import to_float_list


class TestChartsCommon(unittest.TestCase):
    def test_to_float_list_parses_percent_strings(self):
        values = ["9%", " 9 % ", "9,5%", "(10%)"]
        out = to_float_list(values)
        self.assertEqual(out, [9.0, 9.0, 9.5, -10.0])

    def test_to_float_list_parses_ptbr_numbers(self):
        values = ["1.234,56", "0,09", "2.000", ""]
        out = to_float_list(values)
        self.assertEqual(out, [1234.56, 0.09, 2000.0, 0.0])

    def test_to_float_list_rejects_non_numeric(self):
        with self.assertRaises(ValueError):
            to_float_list(["N/A"])  # should still error


if __name__ == "__main__":
    unittest.main()
