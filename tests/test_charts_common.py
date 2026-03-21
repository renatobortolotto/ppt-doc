import unittest

import tempfile
from pathlib import Path

from matplotlib.patches import Wedge
from openpyxl import Workbook

from utils.charts_common import (
    ExcelBarChartSpec,
    close_figure,
    plot_bar_from_excel,
    plot_donut_chart,
    plot_line_from_excel,
    to_float_list,
)


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

    def test_plot_bar_value_scale_and_comma_decimal_labels(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "S"

        ws["A1"].value = 8820
        ws["B1"].value = 9100
        ws["A2"].value = "T1"
        ws["B2"].value = "T2"

        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            xlsx_path = td_path / "t.xlsx"
            out_path = td_path / "out.png"
            wb.save(xlsx_path)

            fig, ax = plot_bar_from_excel(
                ExcelBarChartSpec(
                    file_path=xlsx_path,
                    sheet_name="S",
                    values_range="A1:B1",
                    xlabels_range="A2:B2",
                    output_path=out_path,
                    value_decimals=1,
                    value_scale=0.001,
                    value_decimal_comma=True,
                )
            )
            try:
                labels = [t.get_text() for t in ax.texts]
                self.assertEqual(labels, ["8,8", "9,1"])
            finally:
                close_figure(fig)

    def test_plot_donut_chart_keeps_group_ring_inside_detail_ring(self):
        with tempfile.TemporaryDirectory() as td:
            out_path = Path(td) / "donut.png"

            fig, ax = plot_donut_chart(
                categories=["A", "B", "B"],
                labels=["A1", "B1", "B2"],
                values=[10, 20, 30],
                center_text="Centro",
                output_path=out_path,
            )
            try:
                wedges = [patch for patch in ax.patches if isinstance(patch, Wedge)]
                radii = [round(float(wedge.r), 2) for wedge in wedges]
                self.assertEqual(radii.count(0.74), 2)
                self.assertEqual(radii.count(1.00), 3)
            finally:
                close_figure(fig)


if __name__ == "__main__":
    unittest.main()
