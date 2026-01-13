from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Union

from utils import xlsx_extract


@dataclass(frozen=True)
class TextFieldSpec:
    """Spec for extracting a text value (cell or range) from an XLSX."""

    id: str
    a1_range: str
    sheet: Optional[str] = None


def _coerce_cell_value_to_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


def parse_text_fields_json(path: Union[str, Path]) -> tuple[Optional[str], List[TextFieldSpec]]:
    """Parse a 'text fields' config.

    Accepts either:

    1) Object format (recommended):
        {
          "default_sheet": "DRE Saida",
          "fields": {
            "ROE_RECORRENTE": "K20",
            "OUTRO": {"sheet": "Aba", "cell": "B2"}
          }
        }

    2) List format:
        [
          {"id": "ROE_RECORRENTE", "sheet": "DRE Saida", "cell": "K20"}
        ]

    Returns: (default_sheet, specs)
    """

    path = Path(path)
    raw = json.loads(path.read_text(encoding="utf-8"))

    default_sheet: Optional[str] = None
    specs: List[TextFieldSpec] = []

    if isinstance(raw, list):
        for item in raw:
            if not isinstance(item, dict):
                raise ValueError("Cada item deve ser um objeto")

            field_id = item.get("id") or item.get("ID")
            sheet = item.get("sheet") or item.get("Sheet")
            a1 = item.get("cell") or item.get("Cell") or item.get("range") or item.get("Range")

            if not field_id or not a1:
                raise ValueError("Item precisa ter 'id' e 'cell' (ou 'range')")

            specs.append(TextFieldSpec(id=str(field_id), a1_range=str(a1), sheet=str(sheet) if sheet else None))

        return default_sheet, specs

    if not isinstance(raw, dict):
        raise ValueError("Config deve ser um objeto ou uma lista")

    default_sheet = raw.get("default_sheet") or raw.get("DEFAULT_SHEET")
    fields = raw.get("fields")
    if not isinstance(fields, dict):
        raise ValueError("Config no formato objeto precisa ter 'fields' (objeto)")

    for key, value in fields.items():
        if isinstance(value, str):
            specs.append(TextFieldSpec(id=str(key), a1_range=value, sheet=None))
            continue
        if isinstance(value, dict):
            a1 = value.get("cell") or value.get("range")
            if not a1:
                raise ValueError(f"Campo {key!r} precisa ter 'cell' (ou 'range')")
            sheet = value.get("sheet")
            specs.append(TextFieldSpec(id=str(key), a1_range=str(a1), sheet=str(sheet) if sheet else None))
            continue
        raise ValueError(f"Campo {key!r} inválido: esperado string ou objeto")

    return str(default_sheet) if default_sheet else None, specs


def extract_workbook_text_mapping(
    wb,
    specs: Sequence[TextFieldSpec],
    *,
    default_sheet: Optional[str] = None,
) -> Dict[str, str]:
    out: Dict[str, str] = {}

    for spec in specs:
        sheet_name = spec.sheet or default_sheet
        if not sheet_name:
            raise ValueError(
                f"Spec {spec.id!r} não tem sheet e nenhum default_sheet foi informado"
            )
        if sheet_name not in wb.sheetnames:
            raise ValueError(
                f"Aba não encontrada: {sheet_name!r} (spec={spec.id!r}). Disponíveis: {wb.sheetnames}"
            )

        ws = wb[sheet_name]
        values_2d = xlsx_extract._read_range_2d(ws, spec.a1_range)
        values_1d = xlsx_extract._to_1d(values_2d)
        pieces = [_coerce_cell_value_to_str(v) for v in values_1d]

        # If a range produces multiple cells, join with ", ".
        if not pieces:
            out[spec.id] = ""
        elif len(pieces) == 1:
            out[spec.id] = pieces[0]
        else:
            non_empty = [p for p in pieces if p != ""]
            out[spec.id] = ", ".join(non_empty) if non_empty else ""

    return out


def extract_xlsx_to_text_mapping(
    xlsx_path: Union[str, Path],
    specs: Sequence[TextFieldSpec],
    *,
    default_sheet: Optional[str] = None,
) -> Dict[str, str]:
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"XLSX não encontrado: {xlsx_path}")

    wb = xlsx_extract._load_workbook(filename=xlsx_path, data_only=True)
    out = extract_workbook_text_mapping(wb, specs, default_sheet=default_sheet)

    # Excel formulas: openpyxl does not calculate formulas.
    # If the file was not saved with cached results, data_only=True may return None.
    # For VAR_* fields (quarter deltas), try a fallback read from data_only=False and
    # use the cached value if present.
    var_specs = [s for s in specs if str(s.id).upper().startswith("VAR_")]
    if var_specs and any(out.get(s.id, "") == "" for s in var_specs):
        wb_formula = xlsx_extract._load_workbook(filename=xlsx_path, data_only=False)
        for spec in var_specs:
            if out.get(spec.id, "") != "":
                continue

            sheet_name = spec.sheet or default_sheet
            if not sheet_name or sheet_name not in wb_formula.sheetnames:
                continue

            # Only attempt for single-cell references.
            try:
                min_col, min_row, max_col, max_row = xlsx_extract._range_boundaries(spec.a1_range)
            except Exception:
                continue
            if min_col != max_col or min_row != max_row:
                continue

            ws = wb_formula[sheet_name]
            v = ws.cell(row=min_row, column=min_col).value
            if v is None:
                continue
            # If it's a formula string, we can't evaluate here.
            if isinstance(v, str) and v.strip().startswith("="):
                continue
            out[spec.id] = _coerce_cell_value_to_str(v)

    return out
