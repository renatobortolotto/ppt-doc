import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from utils.xlsx_text_fields import (
    TextFieldSpec,
    extract_workbook_text_mapping_tolerant,
    extract_workbook_text_mapping,
    extract_xlsx_to_text_mapping,
    extract_xlsx_to_text_mapping_tolerant,
    parse_text_fields_json,
)


class _FakeCell:
    def __init__(self, value, *, number_format: str = "General"):
        self.value = value
        self.number_format = number_format


class _FakeWorksheet:
    def __init__(self, values_by_rowcol):
        self._values_by_rowcol = dict(values_by_rowcol)

    def cell(self, row: int, column: int):
        raw = self._values_by_rowcol.get((row, column))
        if isinstance(raw, tuple):
            value, number_format = raw
            return _FakeCell(value, number_format=number_format)
        return _FakeCell(raw)


class _FakeWorkbook:
    def __init__(self, sheets):
        self._sheets = dict(sheets)
        self.sheetnames = list(self._sheets.keys())

    def __getitem__(self, item):
        return self._sheets[item]


def _fake_workbook() -> _FakeWorkbook:
    # Coordinates are (row, col)
    values = {
        # K20 (K=11)
        (20, 11): 0.1234,
        # B2
        (2, 2): "Texto",
        # D4
        (4, 4): 8820,
        # Range C3:D3
        (3, 3): "A",
        (3, 4): "B",
        # E5
        (5, 5): (0.099, "0,0%"),
        # F6
        (6, 6): (0.139, '0,0"p.p."'),
        # G7
        (7, 7): (-0.901, '0,0%;-0,0" p.p."'),
    }
    ws = _FakeWorksheet(values)
    ws2 = _FakeWorksheet({(2, 2): 9.99, (3, 2): 9})
    return _FakeWorkbook({"DRE Saida": ws, "Premissas": ws2})


class TestXlsxTextFields(unittest.TestCase):
    def test_parse_text_fields_object_format(self):
        payload = {
            "default_sheet": "DRE Saida",
            "fields": {
                "ROE_RECORRENTE": "K20",
                "X": {"cell": "B2", "div": 1000, "round": 1, "is_porc": True, "is_pp": False},
            },
        }

        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "text_fields.json"
            path.write_text(json.dumps(payload), encoding="utf-8")

            default_sheet, specs = parse_text_fields_json(path)

        self.assertEqual(default_sheet, "DRE Saida")
        self.assertEqual(len(specs), 2)
        self.assertEqual(specs[0].id, "ROE_RECORRENTE")
        self.assertEqual(specs[0].a1_range, "K20")
        self.assertIsNone(specs[0].div)
        self.assertIsNone(specs[0].round)
        self.assertFalse(specs[0].is_porc)
        self.assertFalse(specs[0].is_pp)
        self.assertEqual(specs[1].div, 1000.0)
        self.assertEqual(specs[1].round, 1)
        self.assertTrue(specs[1].is_porc)
        self.assertFalse(specs[1].is_pp)

    def test_extract_workbook_text_mapping_single_cell_and_range(self):
        specs = [
            TextFieldSpec(id="ROE_RECORRENTE", a1_range="K20", sheet="DRE Saida"),
            TextFieldSpec(id="TEXTO", a1_range="B2", sheet="DRE Saida"),
            TextFieldSpec(id="RANGE", a1_range="C3:D3", sheet="DRE Saida"),
        ]

        out = extract_workbook_text_mapping(_fake_workbook(), specs, default_sheet=None)

        self.assertEqual(out["ROE_RECORRENTE"], "0.1234")
        self.assertEqual(out["TEXTO"], "Texto")
        self.assertEqual(out["RANGE"], "A, B")

    def test_extract_workbook_text_mapping_applies_optional_divisor(self):
        specs = [
            TextFieldSpec(id="EM_MILHARES", a1_range="D4", sheet="DRE Saida", div=1000),
        ]

        out = extract_workbook_text_mapping(_fake_workbook(), specs, default_sheet=None)

        self.assertEqual(out["EM_MILHARES"], "8.82")

    def test_extract_workbook_text_mapping_applies_optional_round_after_divisor(self):
        specs = [
            TextFieldSpec(id="EM_MILHARES_ARRED", a1_range="D4", sheet="DRE Saida", div=1000, round=1),
            TextFieldSpec(id="FIXED_DECIMALS", a1_range="B3", sheet="Premissas", round=2),
        ]

        out = extract_workbook_text_mapping(_fake_workbook(), specs, default_sheet=None)

        self.assertEqual(out["EM_MILHARES_ARRED"], "8.8")
        self.assertEqual(out["FIXED_DECIMALS"], "9.00")

    def test_extract_workbook_text_mapping_preserves_percent_display_when_is_porc(self):
        specs = [
            TextFieldSpec(id="ROE_DISPLAY", a1_range="E5", sheet="DRE Saida", is_porc=True),
            TextFieldSpec(id="ROE_RAW", a1_range="E5", sheet="DRE Saida"),
        ]

        out = extract_workbook_text_mapping(_fake_workbook(), specs, default_sheet=None)

        self.assertEqual(out["ROE_DISPLAY"], "9,9%")
        self.assertEqual(out["ROE_RAW"], "0.099")

    def test_extract_workbook_text_mapping_returns_numeric_value_when_is_pp(self):
        specs = [
            TextFieldSpec(id="PP_NUMERIC", a1_range="F6", sheet="DRE Saida", is_pp=True),
            TextFieldSpec(id="PP_NUMERIC_ROUND", a1_range="F6", sheet="DRE Saida", is_pp=True, round=1),
            TextFieldSpec(id="PP_WINS_OVER_PERCENT", a1_range="G7", sheet="DRE Saida", is_pp=True, is_porc=True, round=1),
            TextFieldSpec(id="PP_RAW", a1_range="F6", sheet="DRE Saida"),
            TextFieldSpec(id="PP_NEG_NUMERIC", a1_range="G7", sheet="DRE Saida", is_pp=True, round=1),
        ]

        out = extract_workbook_text_mapping(_fake_workbook(), specs, default_sheet=None)

        self.assertEqual(out["PP_NUMERIC"], "0.139")
        self.assertEqual(out["PP_NUMERIC_ROUND"], "0.1")
        self.assertEqual(out["PP_WINS_OVER_PERCENT"], "-0.9")
        self.assertEqual(out["PP_RAW"], "0.139")
        self.assertEqual(out["PP_NEG_NUMERIC"], "-0.9")

    def test_extract_workbook_text_mapping_sheet_override(self):
        specs = [
            TextFieldSpec(id="TAXA_DESCONTO", a1_range="B2", sheet="Premissas"),
        ]

        out = extract_workbook_text_mapping(_fake_workbook(), specs, default_sheet="DRE Saida")
        self.assertEqual(out["TAXA_DESCONTO"], "9.99")

    def test_extract_workbook_text_mapping_tolerant_collects_failures(self):
        specs = [
            TextFieldSpec(id="ROE_RECORRENTE", a1_range="K20", sheet="DRE Saida"),
            TextFieldSpec(id="CAMPO_INVALIDO", a1_range="Z99", sheet="Aba Inexistente"),
        ]

        result = extract_workbook_text_mapping_tolerant(
            _fake_workbook(),
            specs,
            default_sheet=None,
        )

        self.assertEqual(result.mapping["ROE_RECORRENTE"], "0.1234")
        self.assertEqual(len(result.failures), 1)
        self.assertEqual(result.failures[0].field_id, "CAMPO_INVALIDO")
        self.assertEqual(result.failures[0].sheet, "Aba Inexistente")
        self.assertEqual(result.failures[0].a1_range, "Z99")

    def test_extract_xlsx_to_text_mapping_uses_load_workbook(self):
        with tempfile.TemporaryDirectory() as td:
            xlsx_path = Path(td) / "file.xlsx"
            xlsx_path.write_bytes(b"placeholder")

            specs = [TextFieldSpec(id="ROE_RECORRENTE", a1_range="K20", sheet="DRE Saida")]

            with patch("utils.xlsx_extract._load_workbook", return_value=_fake_workbook()):
                out = extract_xlsx_to_text_mapping(xlsx_path, specs)

        self.assertEqual(out["ROE_RECORRENTE"], "0.1234")

    def test_extract_xlsx_to_text_mapping_tolerant_uses_load_workbook(self):
        with tempfile.TemporaryDirectory() as td:
            xlsx_path = Path(td) / "file.xlsx"
            xlsx_path.write_bytes(b"placeholder")

            specs = [
                TextFieldSpec(id="ROE_RECORRENTE", a1_range="K20", sheet="DRE Saida"),
                TextFieldSpec(id="CAMPO_INVALIDO", a1_range="B2", sheet="Aba Inexistente"),
            ]

            with patch("utils.xlsx_extract._load_workbook", return_value=_fake_workbook()):
                result = extract_xlsx_to_text_mapping_tolerant(xlsx_path, specs)

        self.assertEqual(result.mapping["ROE_RECORRENTE"], "0.1234")
        self.assertEqual(len(result.failures), 1)
        self.assertEqual(result.failures[0].field_id, "CAMPO_INVALIDO")


if __name__ == "__main__":
    unittest.main()
