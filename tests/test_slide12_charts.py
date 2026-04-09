import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from src.utils.slides.slide12_charts import generate_slide12_charts


class TestSlide12Charts(unittest.TestCase):
    def test_generate_slide12_charts_uses_carteira_ranges(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Carteira"

        for col_offset, label in enumerate(["3T25", "4T25", "1T26"], start=4):
            ws.cell(row=12, column=col_offset).value = label

        ws["C15"] = "Varejo - Produtos de Entrada"
        ws["C22"] = "Varejo Relacional"
        ws["C34"] = "Atacado"

        for col_offset, value in enumerate([10.0, 11.0, 12.0], start=4):
            ws.cell(row=15, column=col_offset).value = value
        for col_offset, value in enumerate([20.0, 21.0, 22.0], start=4):
            ws.cell(row=22, column=col_offset).value = value
        for col_offset, value in enumerate([30.0, 31.0, 32.0], start=4):
            ws.cell(row=34, column=col_offset).value = value

        captured: dict[str, object] = {}

        def _capture_plot(
            *,
            xlabels,
            series_names,
            values,
            output_path,
            bracket_top_gap_scale=0.20,
            bracket_top_gap_min=1.4,
        ):
            captured["xlabels"] = list(xlabels)
            captured["series_names"] = list(series_names)
            captured["values"] = values.tolist()
            captured["output_name"] = Path(output_path).name
            captured["bracket_top_gap_scale"] = bracket_top_gap_scale
            captured["bracket_top_gap_min"] = bracket_top_gap_min

        with tempfile.TemporaryDirectory() as td:
            xlsx_path = Path(td) / "test.xlsx"
            output_dir = Path(td) / "out"
            wb.save(xlsx_path)

            with patch("src.utils.slides.slide12_charts._plot_slide12_stacked", side_effect=_capture_plot):
                files = generate_slide12_charts(xlsx_path=xlsx_path, output_dir=output_dir)

        self.assertEqual([path.name for path in files], ["12_slide12_composicao.png"])
        self.assertEqual(captured["output_name"], "12_slide12_composicao.png")
        self.assertEqual(captured["xlabels"], ["3T25", "4T25", "1T26"])
        self.assertEqual(
            captured["series_names"],
            ["Varejo - Produtos de Entrada", "Varejo Relacional", "Atacado"],
        )
        self.assertEqual(
            captured["values"],
            [
                [0.01, 0.02, 0.03],
                [0.011, 0.021, 0.031],
                [0.012, 0.022, 0.032],
            ],
        )
        self.assertEqual(captured["bracket_top_gap_scale"], 0.12)
        self.assertEqual(captured["bracket_top_gap_min"], 0.9)


if __name__ == "__main__":
    unittest.main()
