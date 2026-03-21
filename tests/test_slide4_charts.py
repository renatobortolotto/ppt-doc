import unittest

from openpyxl import Workbook

from utils.slide4_charts import _build_slide4_center_text, _extract_slide4_donut_series


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
                "Growth",
                "Growth",
                "Growth",
                "Growth",
                "Atacado",
                "Atacado",
                "Atacado",
            ],
        )
        self.assertEqual(
            labels,
            [
                "Veiculos Leves Usados",
                "Outros Veiculos",
                "Paineis Solares",
                "EGV",
                "Cartões",
                "Corporate",
                "Large Corporate + instituicoes financeiras",
                "Pequenas e Medias Empresas (PME)",
            ],
        )
        self.assertEqual(values, [100.0, 120.0, 80.0, 110.0, 90.0, 200.0, 300.0, 400.0])

    def test_build_slide4_center_text_uses_sum_delta_and_placeholder(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Pizza Teste"

        ws["D45"].value = "4T24"

        ws["D47"].value = 40_000
        ws["D48"].value = 30_000
        ws["D49"].value = 20_000

        ws["F47"].value = 44_000
        ws["F48"].value = 33_000
        ws["F49"].value = 22_000

        text = _build_slide4_center_text(ws)

        self.assertEqual(
            text,
            "Carteira\nAmpliada\nR$ 99,0 bi, +10,0% vs 4T24",
        )

    def test_extract_slide4_donut_series_reads_carteira_leaf_rows_and_skips_subtotals(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Carteira"

        rows = {
            14: ("segmento varejo", 69333),
            15: ("Varejo - Produtos de Entrada", 58802),
            16: ("Veiculos Leves Usados", 46888),
            17: ("Veiculos Pesados", 3258),
            18: ("Motos e Veiculos Novos", 4558),
            19: ("Paineis Solares", 3731),
            20: ("Consignado Privado", 333),
            21: ("Outros (FGTS + Saúde)", 35),
            22: ("Varejo Relacional", 10531),
            23: ("Empréstimo com Garantia Vericular (EGV)", 5262),
            24: ("Cartão de Crédito", 5230),
            25: ("CP", 39),
            27: ("Segmento Atacado", 10318),
            28: ("Corporate", 6547),
            29: ("Large Corporate + Instituições Financeiras", 2507),
            30: ("Pequenas e Médias Empresas", 1265),
            34: ("Segmento Atacado Ampliado", 28323),
            35: ("Avais e Fianças Prestados", 6120),
            36: ("TVM privado", 11885),
        }

        for row, (label, value) in rows.items():
            ws.cell(row=row, column=3).value = label
            ws.cell(row=row, column=6).value = value

        categories, labels, values = _extract_slide4_donut_series(ws, source_range="C12:F36")

        self.assertEqual(
            categories,
            [
                "Veiculos Leves",
                "Growth",
                "Growth",
                "Growth",
                "Growth",
                "Atacado",
                "Atacado",
                "Atacado",
            ],
        )
        self.assertEqual(
            labels,
            [
                "Veiculos Leves Usados",
                "Outros Veiculos",
                "Paineis Solares",
                "EGV",
                "Cartões",
                "Corporate",
                "Large Corporate + instituicoes financeiras",
                "Pequenas e Medias Empresas (PME)",
            ],
        )
        self.assertEqual(
            values,
            [46888.0, 7816.0, 3731.0, 5262.0, 5269.0, 12667.0, 14392.0, 1265.0],
        )


if __name__ == "__main__":
    unittest.main()
