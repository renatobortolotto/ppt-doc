from __future__ import annotations

from pathlib import Path

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from utils.charts_common import close_figure


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


def _to_float(v: object) -> float:
    if v is None:
        return float("nan")
    try:
        return float(v)
    except Exception:
        return float("nan")


def _fmt_dec(v: float, decimals: int = 1) -> str:
    return f"{v:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _plot_stacked_bars(
    *,
    xlabels: list[str],
    series_names: list[str],
    values: np.ndarray,
    output_path: Path,
    colors: list[str],
    fmt_label=None,
    bracket_pct: bool = True,
    show_pct_for_n_largest: int = 0,
) -> None:
    """Barra empilhada com labels dentro de cada segmento, total no topo e bracket %."""
    import matplotlib.pyplot as plt
    from matplotlib.colors import to_rgba

    n, m = values.shape
    x = np.arange(n, dtype=float)
    width = 0.62

    if fmt_label is None:
        fmt_label = lambda v: _fmt_dec(v, 1)

    totals = np.nansum(np.where(np.isfinite(values), values, 0.0), axis=1).astype(float)
    avg_by_series = np.nanmean(np.where(np.isfinite(values), values, 0.0), axis=0)
    pct_series_set: set[int] = set()
    if show_pct_for_n_largest > 0:
        pct_series_set = set(int(i) for i in np.argsort(avg_by_series)[-show_pct_for_n_largest:])

    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=240)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    bottom = np.zeros(n, dtype=float)
    segment_centers: list[list[float]] = [[] for _ in range(m)]

    for j in range(m):
        y = np.asarray(values[:, j], dtype=float)
        y_safe = np.where(np.isfinite(y), y, 0.0)
        ax.bar(x, y_safe, width=width, bottom=bottom, color=colors[j % len(colors)], edgecolor="none", zorder=2)

        rgba = to_rgba(colors[j % len(colors)])
        lum = 0.2126 * rgba[0] + 0.7152 * rgba[1] + 0.0722 * rgba[2]
        txt_color = "#ffffff" if lum < 0.50 else "#2f2f2f"

        for i in range(n):
            v = float(y[i])
            if not np.isfinite(v) or abs(v) < 1e-9:
                segment_centers[j].append(float("nan"))
                continue
            yc = float(bottom[i]) + v / 2.0
            segment_centers[j].append(yc)

            if j in pct_series_set and float(totals[i]) > 0:
                pct = v / float(totals[i]) * 100.0
                lbl = f"{fmt_label(v)}\n({pct:.0f}%)"
            else:
                lbl = fmt_label(v)

            ax.text(float(x[i]), yc, lbl, ha="center", va="center", fontsize=9.0, color=txt_color, zorder=4, linespacing=1.3)

        bottom = bottom + y_safe

    totals = bottom.copy()

    total_tops: list[float] = []
    for i, total in enumerate(totals):
        y_lbl = float(total) + max(abs(float(total)) * 0.02, 0.3)
        total_tops.append(y_lbl)
        ax.text(
            float(x[i]),
            y_lbl,
            fmt_label(total),
            ha="center",
            va="bottom",
            fontsize=10.0,
            fontweight="bold" if i == n - 1 else "normal",
            color="#2f2f2f",
            zorder=5,
            clip_on=False,
        )

    if n >= 2:
        abs_max = float(np.nanmax(np.abs(totals)))
        offset_y = max(abs_max * 0.12, 0.5)
        bracket_h = max(abs_max * 0.035, 0.3)
        top_base = max(total_tops) + max(abs_max * 0.22, 1.0)
        max_text_y: float | None = None

        for i in range(1, n):
            prev, curr = float(totals[i - 1]), float(totals[i])
            if not np.isfinite(prev) or not np.isfinite(curr):
                continue
            if bracket_pct:
                if prev == 0:
                    continue
                lbl = f"{(curr / prev - 1.0) * 100.0:+.1f}%".replace(".", ",")
            else:
                lbl = f"{(curr - prev):+.1f} p.p.".replace(".", ",")

            x1, x2 = float(x[i - 1]), float(x[i])
            ax.plot(
                [x1, x1, x2, x2],
                [top_base, top_base + bracket_h, top_base + bracket_h, top_base],
                color="#2f2f2f",
                linewidth=1.2,
                zorder=4,
            )
            ty = top_base + bracket_h + offset_y * 0.25
            ax.text((x1 + x2) / 2.0, ty, lbl, ha="center", va="bottom", fontsize=9.0, color="#2f2f2f", zorder=5)
            max_text_y = ty if max_text_y is None else max(max_text_y, ty)

        if max_text_y is not None:
            ymin, ymax = ax.get_ylim()
            ax.set_ylim(ymin, max(ymax, max_text_y + offset_y * 1.4))

    x_leg = float(x.min()) - 0.95
    for j, name in enumerate(series_names):
        y_ref = next((yc for yc in segment_centers[j] if np.isfinite(yc)), float("nan"))
        if not np.isfinite(y_ref):
            continue
        ax.scatter([x_leg], [y_ref], s=90.0, marker="s", color=colors[j % len(colors)], edgecolors="none", zorder=6)
        ax.text(x_leg + 0.12, y_ref, str(name), ha="left", va="center", fontsize=9.0, color="#2f2f2f", zorder=6, clip_on=False)

    ax.set_xlim(float(x.min()) - 1.25, float(x.max()) + 0.65)
    ax.set_xticks(x)
    ax.set_xticklabels(xlabels, fontsize=10.0)
    ax.tick_params(axis="x", bottom=False, pad=8)
    ax.set_yticks([])
    for s in ("left", "right", "top", "bottom"):
        ax.spines[s].set_visible(False)
    ax.grid(False)
    ax.margins(x=0.05, y=0.12)

    fig.tight_layout(pad=0.2)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=450, transparent=True, bbox_inches="tight", pad_inches=0.05)
    close_figure(fig)


def _plot_double_bars(
    *,
    categories: list[str],
    label_a: str,
    label_b: str,
    values_a: list[float],
    values_b: list[float],
    output_path: Path,
    color_a: str = "#123a7a",
    color_b: str = "#5B8FF9",
    fmt_label=None,
) -> None:
    """Barras verticais duplas (grouped) para comparativo entre dois periodos."""
    import matplotlib.pyplot as plt

    if fmt_label is None:
        fmt_label = lambda v: _fmt_dec(v, 1)

    n = len(categories)
    x = np.arange(n, dtype=float)
    bar_w = 0.35
    gap = 0.03

    fig, ax = plt.subplots(figsize=(12, 5), dpi=240)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    vals_a = np.asarray(values_a, dtype=float)
    vals_b = np.asarray(values_b, dtype=float)

    x_a = x - bar_w / 2 - gap / 2
    x_b = x + bar_w / 2 + gap / 2

    ax.bar(x_a, np.where(np.isfinite(vals_a), vals_a, 0.0), width=bar_w, color=color_a, edgecolor="none", zorder=2)
    ax.bar(x_b, np.where(np.isfinite(vals_b), vals_b, 0.0), width=bar_w, color=color_b, edgecolor="none", zorder=2)

    abs_max = float(np.nanmax(np.abs(np.concatenate([vals_a, vals_b]))))
    y_offset = max(abs_max * 0.015, 0.3)

    for i in range(n):
        va = float(vals_a[i])
        vb = float(vals_b[i])
        if np.isfinite(va):
            ax.text(float(x_a[i]), va + y_offset, fmt_label(va), ha="center", va="bottom", fontsize=7.5, color="#2f2f2f", zorder=4)
        if np.isfinite(vb):
            ax.text(
                float(x_b[i]),
                vb + y_offset,
                fmt_label(vb),
                ha="center",
                va="bottom",
                fontsize=7.5,
                fontweight="bold",
                color="#2f2f2f",
                zorder=4,
            )

    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=7.5, rotation=45, ha="right")
    ax.tick_params(axis="x", bottom=False, pad=4)
    ax.set_yticks([])
    for s in ("left", "right", "top", "bottom"):
        ax.spines[s].set_visible(False)
    ax.grid(False)
    ax.margins(x=0.02, y=0.15)

    from matplotlib.patches import Patch

    legend_handles = [
        Patch(facecolor=color_a, label=label_a),
        Patch(facecolor=color_b, label=label_b),
    ]
    ax.legend(handles=legend_handles, loc="upper center", bbox_to_anchor=(0.5, -0.22), fontsize=9.0, frameon=False, ncol=2)

    fig.tight_layout(pad=0.3)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=450, transparent=True, bbox_inches="tight", pad_inches=0.08)
    close_figure(fig)


def generate_slide20_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
    """Slide 20:
    - A1:D5 -> barra empilhada CIB (3 series x 3 periodos), bracket %.
    - G1:I22 -> barras horizontais duplas comparativo 4T24 vs 4T25 (20 categorias).
    """

    output_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(filename=xlsx_path, data_only=True)
    sheet_name = "slide_20"
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Aba nao encontrada: {sheet_name!r}. Disponiveis: {wb.sheetnames}")
    ws = wb[sheet_name]

    generated: list[Path] = []

    labels_a = [("" if v is None else str(v)).strip() for v in _read_range_row(ws, "B2:D2")]
    series_names_a: list[str] = []
    rows_a: list[list[float]] = []
    for row in [3, 4, 5]:
        raw = _read_range_row(ws, f"A{row}:D{row}")
        series_names_a.append(("" if raw[0] is None else str(raw[0])).strip())
        rows_a.append([_to_float(v) for v in raw[1:]])

    values_a = np.asarray(rows_a, dtype=float).T

    out36 = output_dir / "36_cib_empilhado.png"
    _plot_stacked_bars(
        xlabels=labels_a,
        series_names=series_names_a,
        values=values_a,
        output_path=out36,
        colors=["#123a7a", "#5B8FF9", "#AFC8F5"],
        fmt_label=lambda v: _fmt_dec(v, 1),
        bracket_pct=True,
        show_pct_for_n_largest=2,
    )
    generated.append(out36)

    label_period_a = ("" if ws["H2"].value is None else str(ws["H2"].value)).strip()
    label_period_b = ("" if ws["I2"].value is None else str(ws["I2"].value)).strip()

    categories: list[str] = []
    vals_period_a: list[float] = []
    vals_period_b: list[float] = []
    for row in range(3, 23):
        cat = ws.cell(row=row, column=7).value
        va = ws.cell(row=row, column=8).value
        vb = ws.cell(row=row, column=9).value
        categories.append("" if cat is None else str(cat).strip())
        vals_period_a.append(_to_float(va))
        vals_period_b.append(_to_float(vb))

    out37 = output_dir / "37_carteira_atacado_comparativo.png"
    _plot_double_bars(
        categories=categories,
        label_a=label_period_a,
        label_b=label_period_b,
        values_a=vals_period_a,
        values_b=vals_period_b,
        output_path=out37,
        fmt_label=lambda v: f"{int(round(v))}",
    )
    generated.append(out37)

    return generated


if __name__ == "__main__":
    xlsx = Path("testing.xlsx")
    out = Path(".")
    if xlsx.exists():
        files = generate_slide20_charts(xlsx_path=xlsx, output_dir=out)
        print(f"Gerados: {files}")
    else:
        print(f"Arquivo {xlsx} nao encontrado")
