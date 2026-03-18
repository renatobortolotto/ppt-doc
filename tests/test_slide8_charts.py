import unittest

from openpyxl import Workbook

from utils.slide8_charts import _read_slide8_revenue_rows


class TestSlide8Charts(unittest.TestCase):
    def test_read_slide8_revenue_rows_splits_trimestres_and_9m(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "slide_8"

        labels = ["3T24", "2T25", "3T25", "9M24", "9M25"]
        margem_values = [2372, 2311, 2295, 6782, 6976]
        servicos_values = [685, 556, 624, 1984, 1798]

        for col_offset, label in enumerate(labels, start=4):
            ws.cell(row=5, column=col_offset).value = label

        for col_offset, value in enumerate(margem_values, start=4):
            ws.cell(row=6, column=col_offset).value = value

        for col_offset, value in enumerate(servicos_values, start=4):
            ws.cell(row=9, column=col_offset).value = value

        (tri_labels, tri_series, tri_values), (nm_labels, nm_series, nm_values) = _read_slide8_revenue_rows(
            ws,
            xlabels_range="D5:H5",
            margem_values_range="D6:H6",
            servicos_values_range="D9:H9",
            trimestres_count=3,
        )

        self.assertEqual(tri_labels, ["3T24", "2T25", "3T25"])
        self.assertEqual(tri_series, ["Margem Financeira Bruta", "Serviços e Seguros"])
        self.assertEqual(tri_values.tolist(), [[2372.0, 685.0], [2311.0, 556.0], [2295.0, 624.0]])

        self.assertEqual(nm_labels, ["9M24", "9M25"])
        self.assertEqual(nm_series, ["Margem Financeira Bruta", "Serviços e Seguros"])
        self.assertEqual(nm_values.tolist(), [[6782.0, 1984.0], [6976.0, 1798.0]])


if __name__ == "__main__":
    unittest.main()
