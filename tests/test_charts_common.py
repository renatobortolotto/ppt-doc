import unittest

import tempfile
from pathlib import Path

from openpyxl import Workbook

from utils.charts_common import close_figure, plot_line_from_excel, to_float_list


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

    def test_plot_line_percent_formatted_cells_scale_to_points(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "S"

        # Values stored as fractions but formatted as percent
        ws["A1"].value = 0.09
        ws["A1"].number_format = "0%"
        ws["B1"].value = 0.12
        ws["B1"].number_format = "0%"

        ws["A2"].value = "T1"
        ws["B2"].value = "T2"

        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            xlsx_path = td_path / "t.xlsx"
            out_path = td_path / "out.png"
            wb.save(xlsx_path)

            fig, ax = plot_line_from_excel(
                file_path=xlsx_path,
                sheet_name="S",
                values_range="A1:B1",
                xlabels_range="A2:B2",
                output_path=out_path,
                fmt_as_percent=True,
                smooth=False,
                show_markers=False,
            )
            try:
                labels = [t.get_text() for t in ax.texts]
                self.assertEqual(labels, ["9,0%", "12,0%"])
            finally:
                close_figure(fig)


if __name__ == "__main__":
    unittest.main()
