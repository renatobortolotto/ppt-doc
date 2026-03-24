import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from utils.slide9_charts import _wrap_words, generate_slide9_charts


class TestSlide9Charts(unittest.TestCase):
    def test_wrap_words_breaks_long_legend_labels(self):
        self.assertEqual(
            _wrap_words("Recuperação de Crédito", max_line_len=15),
            "Recuperação de\nCrédito",
        )

    def test_generate_slide9_charts_uses_updated_sources(self):
        wb = Workbook()
        ws_custo = wb.active
        ws_custo.title = "Tabelas"
        ws_cobertura = wb.create_sheet("Qualidade Cart 4966")

        for col_offset, label in enumerate(["9M23", "9M24", "9M25"], start=4):
            ws_custo.cell(row=2, column=col_offset).value = label
        for col_offset, value in enumerate([100, -110, 120], start=4):
            ws_custo.cell(row=13, column=col_offset).value = value
        for col_offset, value in enumerate([10, 12, 14], start=4):
            ws_custo.cell(row=5, column=col_offset).value = value

        for col_offset, label in enumerate(["4T24", "1T25"], start=7):
            ws_custo.cell(row=2, column=col_offset).value = label
        for col_offset, value in enumerate([-40, 45], start=7):
            ws_custo.cell(row=13, column=col_offset).value = value
        for col_offset, value in enumerate([4, 5], start=7):
            ws_custo.cell(row=5, column=col_offset).value = value

        for col_offset, label in enumerate(["4T24", "1T25", "2T25"], start=4):
            ws_cobertura.cell(row=2, column=col_offset).value = label
        for col_offset, value in enumerate([1.6, 1.7, 1.8], start=4):
            ws_cobertura.cell(row=17, column=col_offset).value = value

        captured: dict[str, dict[str, object]] = {}
        captured_lines: dict[str, dict[str, object]] = {}

        def _capture_stacked(*, xlabels, series_names, values, output_path, colors):
            captured[Path(output_path).name] = {
                "xlabels": list(xlabels),
                "series_names": list(series_names),
                "values": values.tolist(),
                "colors": list(colors),
            }

        def _capture_cobertura(*, xlabels, values, output_path, highlight_last_count=3):
            captured[Path(output_path).name] = {
                "xlabels": list(xlabels),
                "values": list(values),
                "highlight_last_count": highlight_last_count,
            }

        def _capture_line(*, file_path, sheet_name, values_range, xlabels_range, output_path, **kwargs):
            captured_lines[Path(output_path).name] = {
                "file_path": str(file_path),
                "sheet_name": sheet_name,
                "values_range": values_range,
                "xlabels_range": xlabels_range,
                "kwargs": dict(kwargs),
            }
            return object(), object()

        with tempfile.TemporaryDirectory() as td:
            xlsx_path = Path(td) / "test.xlsx"
            output_dir = Path(td) / "out"
            wb.save(xlsx_path)

            with patch("utils.slide9_charts._plot_stacked_bars_with_total", side_effect=_capture_stacked):
                with patch("utils.slide9_charts._plot_indice_cobertura_percent", side_effect=_capture_cobertura):
                    with patch("utils.slide9_charts.plot_line_from_excel", side_effect=_capture_line):
                        with patch("utils.slide9_charts.close_figure"):
                            files = generate_slide9_charts(xlsx_path=xlsx_path, output_dir=output_dir)

        self.assertEqual(len(files), 5)
        self.assertEqual(
            captured["09_custo_credito_trimestres.png"]["xlabels"],
            ["4T24", "1T25"],
        )
        self.assertEqual(
            captured["09_custo_credito_trimestres.png"]["series_names"],
            ["PDD Expandida", "Recuperação de Crédito"],
        )
        self.assertEqual(
            captured["09_custo_credito_trimestres.png"]["values"],
            [[40.0, -4.0], [45.0, -5.0]],
        )

        self.assertEqual(
            captured["09_custo_credito_9m.png"]["xlabels"],
            ["9M23", "9M24", "9M25"],
        )
        self.assertEqual(
            captured["09_custo_credito_9m.png"]["values"],
            [[100.0, -10.0], [110.0, -12.0], [120.0, -14.0]],
        )

        self.assertEqual(
            captured["09_indice_cobertura.png"]["xlabels"],
            ["4T24", "1T25", "2T25"],
        )
        self.assertEqual(
            captured["09_indice_cobertura.png"]["values"],
            [1.6, 1.7, 1.8],
        )
        self.assertEqual(captured["09_indice_cobertura.png"]["highlight_last_count"], 2)

        self.assertEqual(
            captured_lines["09_custo_variacao_custo_credito.png"]["sheet_name"],
            "Tabelas",
        )
        self.assertEqual(
            captured_lines["09_custo_variacao_custo_credito.png"]["xlabels_range"],
            "D2:F2",
        )
        self.assertEqual(
            captured_lines["09_custo_variacao_custo_credito.png"]["values_range"],
            "D10:F10",
        )
        self.assertTrue(captured_lines["09_custo_variacao_custo_credito.png"]["kwargs"]["fmt_as_percent"])
        self.assertEqual(captured_lines["09_custo_variacao_custo_credito.png"]["kwargs"]["label_fontsize"], 28.0)
        self.assertEqual(captured_lines["09_custo_variacao_custo_credito.png"]["kwargs"]["marker_size"], 160.0)

        self.assertEqual(
            captured_lines["09_custo_variacao_custo_credito_9m.png"]["sheet_name"],
            "Tabelas",
        )
        self.assertEqual(
            captured_lines["09_custo_variacao_custo_credito_9m.png"]["xlabels_range"],
            "G2:H2",
        )
        self.assertEqual(
            captured_lines["09_custo_variacao_custo_credito_9m.png"]["values_range"],
            "G10:H10",
        )
        self.assertTrue(captured_lines["09_custo_variacao_custo_credito_9m.png"]["kwargs"]["fmt_as_percent"])
        self.assertEqual(captured_lines["09_custo_variacao_custo_credito_9m.png"]["kwargs"]["label_fontsize"], 32.0)
        self.assertEqual(captured_lines["09_custo_variacao_custo_credito_9m.png"]["kwargs"]["marker_size"], 220.0)
        self.assertEqual(captured_lines["09_custo_variacao_custo_credito_9m.png"]["kwargs"]["x_margin"], 0.55)


if __name__ == "__main__":
    unittest.main()
