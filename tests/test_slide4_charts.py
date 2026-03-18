import unittest

from openpyxl import Workbook

from utils.slide4_charts import _extract_slide4_donut_series


class TestSlide4Charts(unittest.TestCase):
    def test_extract_slide4_donut_series_groups_complex_range(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Pizza Teste"

        rows = [
            ("Veiculos Leves Usados", 100.0),
            ("Corporate", 200.0),
            ("Large Corporate + instituições financeiras", 300.0),
            ("Pequenas e Médias Empresas (PME)", 400.0),
            ("Veiculos Pesados e Motos", 50.0),
            ("Veiculos Novos", 70.0),
            ("Paineis Solares", 80.0),
            ("Cartão de Crédito", 90.0),
            ("Empréstimos com Garantia Veicular (EGV)", 110.0),
        ]

        for idx, (label, value) in enumerate(rows, start=12):
            ws.cell(row=idx, column=3).value = label
            ws.cell(row=idx, column=6).value = value

        categories, labels, values = _extract_slide4_donut_series(ws, source_range="C12:F20")

        self.assertEqual(
            categories,
            [
                "Veiculos Leves",
                "Atacado",
                "Atacado",
                "Atacado",
                "Growth",
                "Growth",
                "Growth",
                "Growth",
            ],
        )
        self.assertEqual(
            labels,
            [
                "Veiculos Leves Usados",
                "Corporate",
                "Large Corporate + instituicoes financeiras",
                "Pequenas e Medias Empresas (PME)",
                "Outros Veiculos",
                "Paineis Solares",
                "Cartões",
                "EGV",
            ],
        )
        self.assertEqual(values, [100.0, 200.0, 300.0, 400.0, 120.0, 80.0, 90.0, 110.0])


if __name__ == "__main__":
    unittest.main()
