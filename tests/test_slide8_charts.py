import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from utils.slide8_charts import _read_stacked_rows, _wrap_words, generate_slide8_charts


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

    def test_wrap_words_breaks_long_label(self):
        self.assertEqual(
            _wrap_words("Serviços e Seguros", max_line_len=11),
            "Serviços e\nSeguros",
        )

    def test_generate_slide8_charts_creates_expected_pngs(self):
        wb = Workbook()
        ws_dre = wb.active
        ws_dre.title = "DRE Saida 2"
        ws_tabelas = wb.create_sheet("Tabelas")

        for col_offset, label in enumerate(["3T25", "4T25", "1T26"], start=4):
            ws_dre.cell(row=3, column=col_offset).value = label
        for col_offset, label in enumerate(["9M25", "9M26"], start=7):
            ws_dre.cell(row=3, column=col_offset).value = label

        for col_offset, value in enumerate([2372, 2311, 2295], start=4):
            ws_dre.cell(row=5, column=col_offset).value = value
        for col_offset, value in enumerate([685, 556, 624], start=4):
            ws_dre.cell(row=9, column=col_offset).value = value
        for col_offset, value in enumerate([1800, 1750, 1720], start=4):
            ws_dre.cell(row=6, column=col_offset).value = value
        for col_offset, value in enumerate([572, 561, 575], start=4):
            ws_dre.cell(row=7, column=col_offset).value = value

        for col_offset, value in enumerate([4683, 4606], start=7):
            ws_dre.cell(row=5, column=col_offset).value = value
        for col_offset, value in enumerate([1241, 1180], start=7):
            ws_dre.cell(row=9, column=col_offset).value = value
        for col_offset, value in enumerate([3520, 3440], start=7):
            ws_dre.cell(row=6, column=col_offset).value = value
        for col_offset, value in enumerate([1163, 1166], start=7):
            ws_dre.cell(row=7, column=col_offset).value = value

        for col_offset, label in enumerate(["3T25", "4T25", "1T26"], start=4):
            ws_tabelas.cell(row=16, column=col_offset).value = label
        for col_offset, label in enumerate(["9M25", "9M26"], start=7):
            ws_tabelas.cell(row=16, column=col_offset).value = label

        for col_offset, value in enumerate([110, 120, 130], start=4):
            ws_tabelas.cell(row=19, column=col_offset).value = value
        for col_offset, value in enumerate([75, 80, 90], start=4):
            ws_tabelas.cell(row=28, column=col_offset).value = value
        for col_offset, value in enumerate([330, 360], start=7):
            ws_tabelas.cell(row=19, column=col_offset).value = value
        for col_offset, value in enumerate([220, 240], start=7):
            ws_tabelas.cell(row=28, column=col_offset).value = value

        with tempfile.TemporaryDirectory() as td:
            xlsx_path = Path(td) / "test.xlsx"
            output_dir = Path(td) / "out"
            wb.save(xlsx_path)

            files = generate_slide8_charts(xlsx_path=xlsx_path, output_dir=output_dir)
            self.assertEqual(
                [path.name for path in files],
                [
                    "13_slide8_trimestres.png",
                    "14_slide8_9m.png",
                    "15_margem_financeira_bruta_total_trimestres.png",
                    "16_margem_financeira_bruta_total_9m.png",
                    "17_servicos_corretagem_trimestres.png",
                    "18_servicos_corretagem_9m.png",
                ],
            )
            for file_path in files:
                self.assertTrue(file_path.exists())
                self.assertGreater(file_path.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
