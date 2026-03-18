from __future__ import annotations

import json
import math
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
    div: Optional[float] = None
    round: Optional[int] = None
    is_porc: bool = False
    is_pp: bool = False


@dataclass(frozen=True)
class TextFieldFailure:
    field_id: str
    sheet: Optional[str]
    a1_range: str
    error: str


@dataclass(frozen=True)
class TextFieldExtractionResult:
    mapping: Dict[str, str]
    failures: tuple[TextFieldFailure, ...]


def _coerce_cell_value_to_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


def _parse_divisor(raw_value: Any, *, field_id: str) -> Optional[float]:
    if raw_value is None:
        return None
    try:
        divisor = float(raw_value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Campo {field_id!r} tem div invalido: {raw_value!r}") from exc
    if not math.isfinite(divisor) or divisor == 0:
        raise ValueError(f"Campo {field_id!r} tem div invalido: {raw_value!r}")
    return divisor


def _parse_round_digits(raw_value: Any, *, field_id: str) -> Optional[int]:
    if raw_value is None:
        return None
    if isinstance(raw_value, bool):
        raise ValueError(f"Campo {field_id!r} tem round invalido: {raw_value!r}")
    if isinstance(raw_value, float) and not raw_value.is_integer():
        raise ValueError(f"Campo {field_id!r} tem round invalido: {raw_value!r}")
    try:
        round_digits = int(raw_value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Campo {field_id!r} tem round invalido: {raw_value!r}") from exc
    if round_digits < 0:
        raise ValueError(f"Campo {field_id!r} tem round invalido: {raw_value!r}")
    return round_digits


def _parse_bool_flag(raw_value: Any, *, field_id: str, field_name: str) -> bool:
    if raw_value is None:
        return False
    if isinstance(raw_value, bool):
        return raw_value
    raise ValueError(f"Campo {field_id!r} tem {field_name} invalido: {raw_value!r}")


def _apply_divisor_to_value(value: Any, *, divisor: Optional[float]) -> Any:
    if divisor is None or value is None:
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (datetime, date)):
        return value
    if isinstance(value, (int, float)):
        return float(value) / float(divisor)
    return value


def _apply_round_to_value(value: Any, *, round_digits: Optional[int]) -> Any:
    if round_digits is None or value is None:
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (datetime, date)):
        return value
    if isinstance(value, (int, float)):
        numeric = float(value)
        if not math.isfinite(numeric):
            return value
        return format(numeric, f".{round_digits}f")
    return value


def _format_extracted_value(
    value: Any,
    *,
    divisor: Optional[float],
    round_digits: Optional[int],
) -> str:
    value = _apply_divisor_to_value(value, divisor=divisor)
    value = _apply_round_to_value(value, round_digits=round_digits)
    return _coerce_cell_value_to_str(value)


def _cell_is_percent_formatted(cell) -> bool:
    try:
        fmt = cell.number_format
    except Exception:
        return False
    if not fmt:
        return False
    return "%" in str(fmt)


def _cell_is_pp_formatted(cell) -> bool:
    try:
        fmt = cell.number_format
    except Exception:
        return False
    if not fmt:
        return False
    return "p.p." in str(fmt).lower()


def _select_excel_number_format_section(cell) -> str:
    fmt = str(getattr(cell, "number_format", "") or "")
    if not fmt:
        return ""

    raw_value = getattr(cell, "value", None)
    sections = [section.strip() for section in fmt.split(";")]
    if len(sections) <= 1 or not isinstance(raw_value, (int, float)):
        section = sections[0]
    elif float(raw_value) < 0 and len(sections) >= 2:
        section = sections[1]
    elif float(raw_value) == 0 and len(sections) >= 3:
        section = sections[2]
    else:
        section = sections[0]

    while section.startswith("[") and "]" in section:
        section = section.split("]", 1)[1].lstrip()
    return section


def _extract_excel_number_format_parts(cell) -> tuple[int, str, str]:
    selected_fmt = _select_excel_number_format_section(cell)
    normalized_fmt = selected_fmt.replace('"', "").replace("\\", "")

    first_numeric = -1
    last_numeric = -1
    for idx, ch in enumerate(normalized_fmt):
        if first_numeric < 0:
            if ch in {"0", "#"}:
                first_numeric = idx
                last_numeric = idx
            continue
        if ch in {"0", "#", ".", ","}:
            last_numeric = idx
            continue
        break

    if first_numeric < 0 or last_numeric < 0:
        return 0, ".", ""

    numeric_fmt = normalized_fmt[first_numeric : last_numeric + 1]
    suffix = normalized_fmt[last_numeric + 1 :]

    decimals = 0
    decimal_sep = "."
    last_dot = numeric_fmt.rfind(".")
    last_comma = numeric_fmt.rfind(",")
    sep_pos = max(last_dot, last_comma)
    if sep_pos >= 0:
        decimal_sep = numeric_fmt[sep_pos]
        decimals = sum(ch in {"0", "#"} for ch in numeric_fmt[sep_pos + 1 :])

    return decimals, decimal_sep, suffix


def _format_percent_display(cell) -> str:
    raw_value = getattr(cell, "value", None)
    if raw_value is None:
        return ""
    if isinstance(raw_value, str):
        return raw_value.strip()
    if isinstance(raw_value, bool):
        return str(raw_value)
    if not isinstance(raw_value, (int, float)):
        return _coerce_cell_value_to_str(raw_value)

    if not _cell_is_percent_formatted(cell):
        return _coerce_cell_value_to_str(raw_value)

    decimals, decimal_sep, suffix = _extract_excel_number_format_parts(cell)
    pct_value = float(raw_value) * 100.0
    rendered = format(pct_value, f".{decimals}f")
    if decimal_sep == ",":
        rendered = rendered.replace(".", ",")
    return f"{rendered}{suffix or '%'}"


def _format_pp_display(cell) -> str:
    raw_value = getattr(cell, "value", None)
    if raw_value is None:
        return ""
    if isinstance(raw_value, str):
        return raw_value.strip()
    if isinstance(raw_value, bool):
        return str(raw_value)
    if not isinstance(raw_value, (int, float)):
        return _coerce_cell_value_to_str(raw_value)

    if not _cell_is_pp_formatted(cell):
        return _coerce_cell_value_to_str(raw_value)

    decimals, decimal_sep, suffix = _extract_excel_number_format_parts(cell)
    rendered = format(float(raw_value), f".{decimals}f")
    if decimal_sep == ",":
        rendered = rendered.replace(".", ",")
    return f"{rendered}{suffix or 'p.p.'}"


def _iter_cells_in_range(ws, a1_range: str) -> List[Any]:
    min_col, min_row, max_col, max_row = xlsx_extract._range_boundaries(a1_range)
    out: List[Any] = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            out.append(ws.cell(row=row, column=col))
    return out


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

            specs.append(
                TextFieldSpec(
                    id=str(field_id),
                    a1_range=str(a1),
                    sheet=str(sheet) if sheet else None,
                    div=_parse_divisor(item.get("div"), field_id=str(field_id)),
                    round=_parse_round_digits(item.get("round"), field_id=str(field_id)),
                    is_porc=_parse_bool_flag(item.get("is_porc"), field_id=str(field_id), field_name="is_porc"),
                    is_pp=_parse_bool_flag(item.get("is_pp"), field_id=str(field_id), field_name="is_pp"),
                )
            )

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
            specs.append(
                TextFieldSpec(
                    id=str(key),
                    a1_range=str(a1),
                    sheet=str(sheet) if sheet else None,
                    div=_parse_divisor(value.get("div"), field_id=str(key)),
                    round=_parse_round_digits(value.get("round"), field_id=str(key)),
                    is_porc=_parse_bool_flag(value.get("is_porc"), field_id=str(key), field_name="is_porc"),
                    is_pp=_parse_bool_flag(value.get("is_pp"), field_id=str(key), field_name="is_pp"),
                )
            )
            continue
        raise ValueError(f"Campo {key!r} inválido: esperado string ou objeto")

    return str(default_sheet) if default_sheet else None, specs


def _extract_text_value_for_spec(
    wb,
    spec: TextFieldSpec,
    *,
    default_sheet: Optional[str] = None,
) -> tuple[str, str]:
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
    cells = _iter_cells_in_range(ws, spec.a1_range)
    pieces = [
        (
            _format_percent_display(cell)
            if spec.is_porc
            else _format_pp_display(cell)
            if spec.is_pp
            else _format_extracted_value(
                cell.value,
                divisor=spec.div,
                round_digits=spec.round,
            )
        )
        for cell in cells
    ]

    if not pieces:
        return sheet_name, ""
    if len(pieces) == 1:
        return sheet_name, pieces[0]

    non_empty = [p for p in pieces if p != ""]
    return sheet_name, ", ".join(non_empty) if non_empty else ""


def extract_workbook_text_mapping(
    wb,
    specs: Sequence[TextFieldSpec],
    *,
    default_sheet: Optional[str] = None,
) -> Dict[str, str]:
    out: Dict[str, str] = {}

    for spec in specs:
        _sheet_name, value = _extract_text_value_for_spec(
            wb,
            spec,
            default_sheet=default_sheet,
        )
        out[spec.id] = value

    return out


def extract_workbook_text_mapping_tolerant(
    wb,
    specs: Sequence[TextFieldSpec],
    *,
    default_sheet: Optional[str] = None,
) -> TextFieldExtractionResult:
    out: Dict[str, str] = {}
    failures: List[TextFieldFailure] = []

    for spec in specs:
        try:
            sheet_name, value = _extract_text_value_for_spec(
                wb,
                spec,
                default_sheet=default_sheet,
            )
            out[spec.id] = value
        except Exception as exc:
            failures.append(
                TextFieldFailure(
                    field_id=spec.id,
                    sheet=spec.sheet or default_sheet,
                    a1_range=spec.a1_range,
                    error=str(exc),
                )
            )

    return TextFieldExtractionResult(
        mapping=out,
        failures=tuple(failures),
    )


def _apply_var_formula_fallback(
    wb_formula,
    specs: Sequence[TextFieldSpec],
    out: Dict[str, str],
    *,
    default_sheet: Optional[str] = None,
) -> None:
    var_specs = [
        s
        for s in specs
        if str(s.id).upper().startswith("VAR_") and s.id in out and out.get(s.id, "") == ""
    ]
    for spec in var_specs:
        sheet_name = spec.sheet or default_sheet
        if not sheet_name or sheet_name not in wb_formula.sheetnames:
            continue

        try:
            min_col, min_row, max_col, max_row = xlsx_extract._range_boundaries(spec.a1_range)
        except Exception:
            continue
        if min_col != max_col or min_row != max_row:
            continue

        ws = wb_formula[sheet_name]
        v = ws.cell(row=min_row, column=min_col).value
        cell = ws.cell(row=min_row, column=min_col)
        if v is None:
            continue
        if isinstance(v, str) and v.strip().startswith("="):
            continue
        out[spec.id] = (
            _format_percent_display(cell)
            if spec.is_porc
            else _format_pp_display(cell)
            if spec.is_pp
            else _format_extracted_value(
                v,
                divisor=spec.div,
                round_digits=spec.round,
            )
        )


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
        _apply_var_formula_fallback(
            wb_formula,
            specs,
            out,
            default_sheet=default_sheet,
        )

    return out


def extract_xlsx_to_text_mapping_tolerant(
    xlsx_path: Union[str, Path],
    specs: Sequence[TextFieldSpec],
    *,
    default_sheet: Optional[str] = None,
) -> TextFieldExtractionResult:
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"XLSX não encontrado: {xlsx_path}")

    wb = xlsx_extract._load_workbook(filename=xlsx_path, data_only=True)
    result = extract_workbook_text_mapping_tolerant(
        wb,
        specs,
        default_sheet=default_sheet,
    )
    out = dict(result.mapping)

    var_specs = [s for s in specs if str(s.id).upper().startswith("VAR_")]
    if var_specs and any(out.get(s.id, "") == "" for s in var_specs):
        wb_formula = xlsx_extract._load_workbook(filename=xlsx_path, data_only=False)
        _apply_var_formula_fallback(
            wb_formula,
            specs,
            out,
            default_sheet=default_sheet,
        )

    return TextFieldExtractionResult(
        mapping=out,
        failures=result.failures,
    )
