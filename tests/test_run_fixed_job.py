import unittest

from run_fixed_job import _parse_only_slides


class TestRunFixedJob(unittest.TestCase):
    def test_parse_only_slides_accepts_ranges_and_list(self):
        self.assertEqual(_parse_only_slides("1-3, 8, 10"), (1, 2, 3, 8, 10))

    def test_parse_only_slides_rejects_invalid_values(self):
        with self.assertRaises(ValueError):
            _parse_only_slides("3-a")


if __name__ == "__main__":
    unittest.main()
