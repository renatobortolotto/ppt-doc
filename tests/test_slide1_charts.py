import unittest
from pathlib import Path

from utils.slide1_charts import generate_slide1_charts


class TestSlide1Charts(unittest.TestCase):
    def test_legacy_slide1_alias_raises_clear_error(self):
        with self.assertRaisesRegex(RuntimeError, "foi removido do workspace atual"):
            generate_slide1_charts(
                xlsx_path=Path("testing.xlsx"),
                output_dir=Path("."),
            )


if __name__ == "__main__":
    unittest.main()
