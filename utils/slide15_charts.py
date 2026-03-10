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


def _to_float_or_nan(v: object) -> float:
    if v is None:
        return float("nan")
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return float("nan")
        s = s.replace("%", "").replace(" ", "")
        s = s.replace(".", "").replace(",", ".") if ("," in s and "." in s) else s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return float("nan")
    try:
        return float(v)
    except Exception:
        return float("nan")


def _fmt_pct(v: float, decimals: int = 1) -> str:
    return f"{float(v):.{int(decimals)}f}%".replace(".", ",")


def _plot_stacked_percent(
    *,
    xlabels: list[str],
    series_names: list[str],
    values: np.ndarray,  # [n_bars, n_series], can contain NaN
    output_path: Path,
    bracket_mode: str,  # "pct_consecutive" | "pp_first_last"
) -> None:
    import matplotlib.pyplot as plt
    from matplotlib.colors import to_rgba

    n, m = values.shape
    if n == 0 or m == 0:
        raise ValueError("Sem dados para plotar")

    colors = ["#123a7a", "#5B8FF9", "#AFC8F5"]

    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=240)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    x = np.arange(n, dtype=float)
    width = 0.62
    bottom = np.zeros(n, dtype=float)
    centers: list[list[float]] = [[] for _ in range(m)]

    for j in range(m):
        y_raw = np.asarray(values[:, j], dtype=float)
        y = np.where(np.isfinite(y_raw), y_raw, 0.0)
        ax.bar(x, y, width=width, bottom=bottom, color=colors[j % len(colors)], edgecolor="none", zorder=2)

        rgba = to_rgba(colors[j % len(colors)])
        lum = 0.2126 * rgba[0] + 0.7152 * rgba[1] + 0.0722 * rgba[2]
        txt_color = "#ffffff" if lum < 0.50 else "#2f2f2f"

        for i in range(n):
            v = float(y_raw[i])
            if not np.isfinite(v) or abs(v) < 1e-12:
                centers[j].append(float("nan"))
                continue
            yc = float(bottom[i]) + float(y[i]) / 2.0
            centers[j].append(yc)
            ax.text(float(x[i]), yc, _fmt_pct(v * 100.0, 1), ha="center", va="center", fontsize=8.7, color=txt_color, zorder=4)

        bottom = bottom + y

    totals = np.nansum(np.where(np.isfinite(values), values, 0.0), axis=1).astype(float)

    total_label_tops: list[float] = []
    for i, total in enumerate(totals):
        y_lbl = float(total) + max(abs(float(total)) * 0.03, 0.003)
        total_label_tops.append(y_lbl)
        ax.text(
            float(x[i]),
            y_lbl,
            _fmt_pct(total * 100.0, 1),
            ha="center",
            va="bottom",
            fontsize=10.0,
            fontweight="bold" if i == n - 1 else "normal",
            color="#2f2f2f",
            zorder=5,
            clip_on=False,
        )

    # Brackets
    if n >= 2:
        abs_max = float(np.nanmax(np.abs(totals))) if np.isfinite(np.nanmax(np.abs(totals))) else 0.0
        offset_y = max(abs_max * 0.10, 0.008)
        bracket_h = max(abs_max * 0.03, 0.006)
        top_base = max(total_label_tops) + max(abs_max * 0.22, 0.012)
        max_text_y: float | None = None

        pairs: list[tuple[int, int]]
        if bracket_mode == "pp_first_last":
            pairs = [(0, n - 1)]
        else:  # pct_consecutive | pp_consecutive
            pairs = [(i - 1, i) for i in range(1, n)]

        for p, c in pairs:
            prev = float(totals[p])
            curr = float(totals[c])
            if not np.isfinite(prev) or not np.isfinite(curr):
                continue

            x1, x2 = float(x[p]), float(x[c])
            y = top_base
            ax.plot([x1, x1, x2, x2], [y, y + bracket_h, y + bracket_h, y], color="#2f2f2f", linewidth=1.2, zorder=4)

            if bracket_mode in ("pp_first_last", "pp_consecutive"):
                delta_pp = (curr - prev) * 100.0
                lbl = f"{delta_pp:+.1f} p.p.".replace(".", ",")
            else:
                if prev == 0:
                    continue
                pct = (curr / prev - 1.0) * 100.0
                lbl = f"{pct:+.1f}%".replace(".", ",")

            ty = y + bracket_h + offset_y * 0.25
            ax.text((x1 + x2) / 2.0, ty, lbl, ha="center", va="bottom", fontsize=9.0, color="#2f2f2f", zorder=5)
            max_text_y = ty if max_text_y is None else max(max_text_y, ty)

        if max_text_y is not None:
            ymin, ymax = ax.get_ylim()
            ax.set_ylim(ymin, max(ymax, max_text_y + offset_y * 1.2))

    # Inline legend on left – with word-wrap to avoid overlapping the chart
    import textwrap

    _WRAP_WIDTH = 13  # chars per line
    max_lines = max(
        (len(textwrap.wrap(str(name), width=_WRAP_WIDTH)) or 1) for name in series_names
    ) if series_names else 1
    # Each extra line needs ~0.07 extra x-units of breathing room; base margin = 1.2
    left_margin = 1.2 + max(0.0, (max_lines - 1) * 0.07)
    x_leg = float(x.min()) - left_margin + 0.3

    for j, name in enumerate(series_names):
        y_ref = centers[j][-1] if centers[j] else float("nan")
        if not np.isfinite(y_ref):
            for yc in centers[j]:
                if np.isfinite(yc):
                    y_ref = yc
                    break
        if not np.isfinite(y_ref):
            continue
        name_wrapped = textwrap.fill(str(name), width=_WRAP_WIDTH)
        ax.scatter([x_leg], [float(y_ref)], s=90.0, marker="s", color=colors[j % len(colors)], edgecolors="none", zorder=6)
        ax.text(x_leg + 0.12, float(y_ref), name_wrapped, ha="left", va="center", fontsize=8.8, color="#2f2f2f", zorder=6, clip_on=False)

    ax.set_xlim(float(x.min()) - (left_margin + 0.3), float(x.max()) + 0.65)
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


def _plot_bridge_chart(
    *,
    xlabels: list[str],
    series_names: list[str],
    values: np.ndarray,  # [n_bars, n_series]
    output_path: Path,
) -> None:
    """Bridge/waterfall chart:
    - First and last bars: fully stacked, smallest series on top (largest at bottom).
    - Middle bars: floating projection boxes showing incremental deltas.
    - pp bracket comparing first and last bar totals.
    """
    import textwrap
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    from matplotlib.colors import to_rgba

    n, m = values.shape
    if n == 0 or m == 0:
        raise ValueError("Sem dados para plotar")

    colors = ["#123a7a", "#5B8FF9", "#AFC8F5"]
    bridge_color = "#5B8FF9"
    bridge_edge = "#2c5faa"

    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=240)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    x = np.arange(n, dtype=float)
    width = 0.52

    full_cols = [0, n - 1]
    bridge_cols = list(range(1, n - 1))

    # Stack order: largest average at bottom, smallest on top
    avg_vals = np.nanmean(
        np.where(np.isfinite(values[np.ix_(full_cols, list(range(m)))]), values[np.ix_(full_cols, list(range(m)))], 0.0),
        axis=0,
    )
    stack_order = list(np.argsort(avg_vals)[::-1])  # largest → bottom

    full_bar_totals: dict[int, float] = {}
    series_centers: list[tuple[int, float]] = []  # (series_idx, y_center) for legend

    # ---- Full stacked bars ----
    for i in full_cols:
        bottom_i = 0.0
        bar_centers: list[tuple[int, float]] = []
        for j in stack_order:
            v_raw = values[i, j]
            v = float(v_raw) if np.isfinite(v_raw) else 0.0
            if abs(v) < 1e-12:
                continue
            color = colors[j % len(colors)]
            ax.bar(x[i], v, width=width, bottom=bottom_i, color=color, edgecolor="none", zorder=2)

            rgba = to_rgba(color)
            lum = 0.2126 * rgba[0] + 0.7152 * rgba[1] + 0.0722 * rgba[2]
            txt_color = "#ffffff" if lum < 0.50 else "#2f2f2f"
            yc = bottom_i + v / 2.0
            bar_centers.append((j, yc))
            ax.text(float(x[i]), yc, _fmt_pct(v * 100.0, 1),
                    ha="center", va="center", fontsize=8.7, color=txt_color, zorder=4)
            bottom_i += v

        full_bar_totals[i] = bottom_i
        ax.text(float(x[i]), bottom_i + 0.002, _fmt_pct(bottom_i * 100.0, 1),
                ha="center", va="bottom", fontsize=10.0,
                fontweight="bold" if i == n - 1 else "normal",
                color="#2f2f2f", zorder=5, clip_on=False)

        if i == 0:
            series_centers = bar_centers

    # ---- Bridge (floating projection boxes) ----
    cum = full_bar_totals[0]
    prev_x_right = float(x[0]) + width / 2.0

    for i in bridge_cols:
        # Bridge delta comes from first series row (j=0), which holds the delta values
        v_raw = values[i, 0]
        delta = float(v_raw) if np.isfinite(v_raw) else 0.0

        cur_x_left = float(x[i]) - width / 2.0
        cur_x_right = float(x[i]) + width / 2.0

        # Dashed connector at current cumulative level
        ax.plot([prev_x_right + 0.04, cur_x_left - 0.04], [cum, cum],
                color="#bbbbbb", linewidth=0.9, linestyle="--", zorder=1)

        box_h = abs(delta)
        if box_h > 1e-12:
            rect = mpatches.FancyBboxPatch(
                (cur_x_left, cum), width, box_h,
                boxstyle="square,pad=0",
                facecolor=bridge_color, edgecolor=bridge_edge, linewidth=1.8, zorder=3,
            )
            ax.add_patch(rect)
            label_y = cum + box_h / 2.0
            lbl_va = "center"
        else:
            label_y = cum + 0.001
            lbl_va = "bottom"

        ax.text(float(x[i]), label_y, _fmt_pct(delta * 100.0, 1),
                ha="center", va=lbl_va,
                fontsize=9.5, fontweight="bold", color="#2f2f2f", zorder=5)

        cum += delta
        prev_x_right = cur_x_right

    # Dashed connector from last bridge to final bar
    ax.plot([prev_x_right + 0.04, float(x[n - 1]) - width / 2.0 - 0.04], [cum, cum],
            color="#bbbbbb", linewidth=0.9, linestyle="--", zorder=1)

    # ---- pp_first_last bracket ----
    total_first = full_bar_totals[0]
    total_last = full_bar_totals[n - 1]
    abs_max = max(abs(total_first), abs(total_last), abs(cum))
    offset_y = max(abs_max * 0.10, 0.008)
    bracket_h = max(abs_max * 0.03, 0.006)
    top_base = max(list(full_bar_totals.values()) + [cum]) + max(abs_max * 0.22, 0.012)

    x1, x2 = float(x[0]), float(x[n - 1])
    ax.plot([x1, x1, x2, x2], [top_base, top_base + bracket_h, top_base + bracket_h, top_base],
            color="#2f2f2f", linewidth=1.2, zorder=4)
    delta_pp = (total_last - total_first) * 100.0
    lbl = f"{delta_pp:+.1f} p.p.".replace(".", ",")
    ty = top_base + bracket_h + offset_y * 0.25
    ax.text((x1 + x2) / 2.0, ty, lbl, ha="center", va="bottom",
            fontsize=9.0, color="#2f2f2f", zorder=5)

    ymin, ymax = ax.get_ylim()
    ax.set_ylim(ymin, max(ymax, ty + offset_y * 1.2))

    # ---- Inline legend (left side, word-wrapped) ----
    _WRAP_WIDTH = 13
    max_lines = max(
        (len(textwrap.wrap(str(series_names[j]), width=_WRAP_WIDTH)) or 1) for j, _ in series_centers
    ) if series_centers else 1
    left_margin = 1.2 + max(0.0, (max_lines - 1) * 0.07)
    x_leg = float(x.min()) - left_margin + 0.3

    for j, yc in series_centers:
        if j >= len(series_names):
            continue
        name_wrapped = textwrap.fill(str(series_names[j]), width=_WRAP_WIDTH)
        ax.scatter([x_leg], [yc], s=90.0, marker="s",
                   color=colors[j % len(colors)], edgecolors="none", zorder=6)
        ax.text(x_leg + 0.12, yc, name_wrapped, ha="left", va="center",
                fontsize=8.8, color="#2f2f2f", zorder=6, clip_on=False)

    ax.set_xlim(float(x.min()) - (left_margin + 0.3), float(x.max()) + 0.65)
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


def generate_slide15_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
    """Slide 15: B2:G6 normal + J2:O6 com vazios e bracket único em p.p."""

    output_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(filename=xlsx_path, data_only=True)
    sheet_name = "slide_15"
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Aba não encontrada: {sheet_name!r}. Disponíveis: {wb.sheetnames}")
    ws = wb[sheet_name]

    # Block D2:G6
    labels_left = [("" if v is None else str(v)).strip() for v in _read_range_row(ws, "E3:G3")]
    series_left = [("" if v is None else str(v)).strip() for v in _read_range_col(ws, "D4:D6")]
    left_rows: list[list[float]] = []
    for r in range(4, 7):
        left_rows.append([_to_float_or_nan(v) for v in _read_range_row(ws, f"E{r}:G{r}")])
    values_left = np.asarray(left_rows, dtype=float).T  # [3,3]

    out31 = output_dir / "31_indice_basileia_trimestres.png"
    _plot_stacked_percent(
        xlabels=labels_left,
        series_names=series_left,
        values=values_left,
        output_path=out31,
        bracket_mode="pp_consecutive",
    )

    # Block K3:O6
    labels_right = [("" if v is None else str(v)).strip() for v in _read_range_row(ws, "K3:O3")]
    series_right = [("" if v is None else str(v)).strip() for v in _read_range_col(ws, "J4:J6")]
    right_rows: list[list[float]] = []
    for r in range(4, 7):
        right_rows.append([_to_float_or_nan(v) for v in _read_range_row(ws, f"K{r}:O{r}")])
    values_right = np.asarray(right_rows, dtype=float).T  # [5,3], with NaN gaps

    out32 = output_dir / "32_basileia_pp_bridge.png"
    _plot_bridge_chart(
        xlabels=labels_right,
        series_names=series_right,
        values=values_right,
        output_path=out32,
    )

    return [out31, out32]


if __name__ == "__main__":
    xlsx = Path("testing.xlsx")
    out = Path(".")
    if xlsx.exists():
        files = generate_slide15_charts(xlsx_path=xlsx, output_dir=out)
        print(f"Gerados: {files}")
    else:
        print(f"Arquivo {xlsx} não encontrado")
