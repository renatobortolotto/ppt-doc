import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from utils.slide11_charts import _read_named_series_rows, _wrap_words, generate_slide11_charts


class TestSlide11Charts(unittest.TestCase):
    def test_wrap_words_breaks_depreciacao_label(self):
        self.assertEqual(
            _wrap_words("Depreciação e Amortização", max_line_len=16),
            "Depreciação e\nAmortização",
        )

    def test_read_named_series_rows_rejects_blank_ranges(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Tabelas"

        with self.assertRaisesRegex(ValueError, "Range sem labels"):
            _read_named_series_rows(
                ws,
                xlabels_range="D33:F33",
                series_specs=[
                    ("Pessoal", "D35:F35"),
                    ("Administrativas", "D39:F39"),
                ],
            )

    def test_generate_slide11_charts_uses_updated_expense_sources(self):
        wb = Workbook()
        ws_expenses = wb.active
        ws_expenses.title = "Tabelas"
        ws_index = wb.create_sheet("slide_11")

        for col_offset, label in enumerate(["1T25", "2T25", "3T25"], start=4):
            ws_expenses.cell(row=33, column=col_offset).value = label
        for col_offset, value in enumerate([100, 110, 120], start=4):
            ws_expenses.cell(row=35, column=col_offset).value = value
        for col_offset, value in enumerate([40, 42, 44], start=4):
            ws_expenses.cell(row=39, column=col_offset).value = value
        for col_offset, value in enumerate([10, 11, 12], start=4):
            ws_expenses.cell(row=45, column=col_offset).value = value

        for col_offset, label in enumerate(["9M24", "9M25"], start=7):
            ws_expenses.cell(row=33, column=col_offset).value = label
        for col_offset, value in enumerate([330, 360], start=7):
            ws_expenses.cell(row=35, column=col_offset).value = value
        for col_offset, value in enumerate([126, 132], start=7):
            ws_expenses.cell(row=39, column=col_offset).value = value
        for col_offset, value in enumerate([31, 33], start=7):
            ws_expenses.cell(row=45, column=col_offset).value = value

        for col_offset, label in enumerate(["1T25", "2T25", "3T25", "9M24", "9M25"], start=11):
            ws_index.cell(row=3, column=col_offset).value = label
        for col_offset, value in enumerate([0.37, 0.38, 0.39, 0.40, 0.41], start=11):
            ws_index.cell(row=4, column=col_offset).value = value

        captured: dict[str, dict[str, object]] = {}

        def _capture_stacked(*, xlabels, series_names, values, output_path):
            captured[Path(output_path).name] = {
                "xlabels": list(xlabels),
                "series_names": list(series_names),
                "values": values.tolist(),
            }

        def _capture_index(*, xlabels, values, output_path):
            captured[Path(output_path).name] = {
                "xlabels": list(xlabels),
                "values": list(values),
            }

        with tempfile.TemporaryDirectory() as td:
            xlsx_path = Path(td) / "test.xlsx"
            output_dir = Path(td) / "out"
            wb.save(xlsx_path)

            with patch("utils.slide11_charts._plot_stacked_expenses", side_effect=_capture_stacked):
                with patch("utils.slide11_charts._plot_efficiency_index", side_effect=_capture_index):
                    files = generate_slide11_charts(xlsx_path=xlsx_path, output_dir=output_dir)

        self.assertEqual(len(files), 3)
        self.assertEqual(
            captured["22_despesas_pessoal_adm_trimestres.png"]["xlabels"],
            ["1T25", "2T25", "3T25"],
        )
        self.assertEqual(
            captured["22_despesas_pessoal_adm_trimestres.png"]["series_names"],
            ["Pessoal", "Administrativas", "Depreciação e Amortização"],
        )
        self.assertEqual(
            captured["22_despesas_pessoal_adm_trimestres.png"]["values"],
            [[100.0, 40.0, 10.0], [110.0, 42.0, 11.0], [120.0, 44.0, 12.0]],
        )

        self.assertEqual(
            captured["23_despesas_pessoal_adm_9m.png"]["xlabels"],
            ["9M24", "9M25"],
        )
        self.assertEqual(
            captured["23_despesas_pessoal_adm_9m.png"]["values"],
            [[330.0, 126.0, 31.0], [360.0, 132.0, 33.0]],
        )

        self.assertEqual(
            captured["24_indice_eficiencia.png"]["xlabels"],
            ["1T25", "2T25", "3T25", "9M24", "9M25"],
        )
        self.assertEqual(
            captured["24_indice_eficiencia.png"]["values"],
            [0.37, 0.38, 0.39, 0.40, 0.41],
        )


if __name__ == "__main__":
    unittest.main()
