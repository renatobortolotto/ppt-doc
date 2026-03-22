import unittest

from openpyxl import Workbook

from utils.slide8_charts import _read_stacked_rows


class TestSlide8Charts(unittest.TestCase):
    def test_read_stacked_rows_reads_named_series_from_linear_ranges(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "DRE Saida 2"

        labels = ["3T24", "2T25", "3T25"]
        margem_values = [2372, 2311, 2295]
        servicos_values = [685, 556, 624]

        for col_offset, label in enumerate(labels, start=4):
            ws.cell(row=3, column=col_offset).value = label

        for col_offset, value in enumerate(margem_values, start=4):
            ws.cell(row=5, column=col_offset).value = value

        for col_offset, value in enumerate(servicos_values, start=4):
            ws.cell(row=9, column=col_offset).value = value

        xlabels, series, values = _read_stacked_rows(
            ws,
            xlabels_range="D3:F3",
            series_specs=[
                ("Margem Financeira Bruta", "D5:F5"),
                ("Serviços e Seguros", "D9:F9"),
            ],
        )

        self.assertEqual(xlabels, ["3T24", "2T25", "3T25"])
        self.assertEqual(series, ["Margem Financeira Bruta", "Serviços e Seguros"])
        self.assertEqual(values.tolist(), [[2372.0, 685.0], [2311.0, 556.0], [2295.0, 624.0]])


if __name__ == "__main__":
    unittest.main()
