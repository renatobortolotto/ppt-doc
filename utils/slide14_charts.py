from __future__ import annotations

from pathlib import Path

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from utils.charts_common import to_float_list
from utils.slide12_charts import _plot_slide12_stacked


def _read_range_row(ws, cell_range: str) -> list[object]:
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    out: list[object] = []
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            out.append(ws.cell(row=r, column=c).value)
    return out


def _read_range_col(ws, cell_range: str) -> list[object]:
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    out: list[object] = []
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            out.append(ws.cell(row=r, column=c).value)
    return out


def generate_slide14_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
    """Slide 14: composição B2:E10 com a mesma paleta/estilo do slide 12."""

    output_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(filename=xlsx_path, data_only=True)
    sheet_name = "slide_14"
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Aba não encontrada: {sheet_name!r}. Disponíveis: {wb.sheetnames}")
    ws = wb[sheet_name]

    xlabels = [("" if v is None else str(v)).strip() for v in _read_range_row(ws, "C2:E2")]
    series_names = [("" if v is None else str(v)).strip() for v in _read_range_col(ws, "B3:B10")]

    raw_rows: list[list[float]] = []
    for r in range(3, 11):
        raw_rows.append(to_float_list(_read_range_row(ws, f"C{r}:E{r}")))
    values = np.asarray(raw_rows, dtype=float).T  # [n_bars, n_series]

    out30 = output_dir / "30_slide14_composicao.png"
    _plot_slide12_stacked(
        xlabels=xlabels,
        series_names=series_names,
        values=values,
        output_path=out30,
    )
    return [out30]


if __name__ == "__main__":
    xlsx = Path("testing.xlsx")
    out = Path(".")
    if xlsx.exists():
        files = generate_slide14_charts(xlsx_path=xlsx, output_dir=out)
        print(f"Gerados: {files}")
    else:
        print(f"Arquivo {xlsx} não encontrado")
