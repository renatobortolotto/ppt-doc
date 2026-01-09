import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from utils.xlsx_text_fields import (
    TextFieldSpec,
    extract_workbook_text_mapping,
    extract_xlsx_to_text_mapping,
    parse_text_fields_json,
)


class _FakeCell:
    def __init__(self, value):
        self.value = value


class _FakeWorksheet:
    def __init__(self, values_by_rowcol):
        self._values_by_rowcol = dict(values_by_rowcol)

    def cell(self, row: int, column: int):
        return _FakeCell(self._values_by_rowcol.get((row, column)))


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
        # Range C3:D3
        (3, 3): "A",
        (3, 4): "B",
    }
    ws = _FakeWorksheet(values)
    ws2 = _FakeWorksheet({(2, 2): 9.99})
    return _FakeWorkbook({"DRE Saida": ws, "Premissas": ws2})


class TestXlsxTextFields(unittest.TestCase):
    def test_parse_text_fields_object_format(self):
        payload = {
            "default_sheet": "DRE Saida",
            "fields": {"ROE_RECORRENTE": "K20", "X": {"cell": "B2"}},
        }

        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "text_fields.json"
            path.write_text(json.dumps(payload), encoding="utf-8")

            default_sheet, specs = parse_text_fields_json(path)

        self.assertEqual(default_sheet, "DRE Saida")
        self.assertEqual(len(specs), 2)
        self.assertEqual(specs[0].id, "ROE_RECORRENTE")
        self.assertEqual(specs[0].a1_range, "K20")

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

    def test_extract_workbook_text_mapping_sheet_override(self):
        specs = [
            TextFieldSpec(id="TAXA_DESCONTO", a1_range="B2", sheet="Premissas"),
        ]

        out = extract_workbook_text_mapping(_fake_workbook(), specs, default_sheet="DRE Saida")
        self.assertEqual(out["TAXA_DESCONTO"], "9.99")

    def test_extract_xlsx_to_text_mapping_uses_load_workbook(self):
        with tempfile.TemporaryDirectory() as td:
            xlsx_path = Path(td) / "file.xlsx"
            xlsx_path.write_bytes(b"placeholder")

            specs = [TextFieldSpec(id="ROE_RECORRENTE", a1_range="K20", sheet="DRE Saida")]

            with patch("utils.xlsx_extract._load_workbook", return_value=_fake_workbook()):
                out = extract_xlsx_to_text_mapping(xlsx_path, specs)

        self.assertEqual(out["ROE_RECORRENTE"], "0.1234")


if __name__ == "__main__":
    unittest.main()
