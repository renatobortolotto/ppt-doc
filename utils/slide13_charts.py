from __future__ import annotations

from pathlib import Path

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from utils.charts_common import close_figure, to_float_list


def _read_range_row(ws, cell_range: str) -> list[object]:
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    out: list[object] = []
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            out.append(ws.cell(row=r, column=c).value)
    return out


def _fmt_num(v: float) -> str:
    return f"{float(v):.1f}".replace(".", ",")


def _fmt_pct(v: float) -> str:
    return f"{float(v):.0f}%"


def _plot_stacked_originacao(
    *,
    xlabels: list[str],
    series_names: list[str],
    values: np.ndarray,  # [n_bars, 2]
    output_path: Path,
) -> None:
    import matplotlib.pyplot as plt
    from matplotlib.colors import to_rgba

    n, m = values.shape
    if m != 2:
        raise ValueError("Originação deve ter 2 séries")

    colors = ["#123a7a", "#5B8FF9"]  # only blues

    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=240)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    x = np.arange(n, dtype=float)
    width = 0.62
    bottom = np.zeros(n, dtype=float)
    centers: list[list[float]] = [[] for _ in range(m)]

    for j in range(m):
        y = np.asarray(values[:, j], dtype=float)
        ax.bar(x, y, width=width, bottom=bottom, color=colors[j], edgecolor="none", zorder=2)
        rgba = to_rgba(colors[j])
        lum = 0.2126 * rgba[0] + 0.7152 * rgba[1] + 0.0722 * rgba[2]
        txt_color = "#ffffff" if lum < 0.50 else "#2f2f2f"
        for i in range(n):
            v = float(y[i])
            if not np.isfinite(v) or abs(v) < 1e-12:
                centers[j].append(float("nan"))
                continue
            yc = float(bottom[i]) + v / 2.0
            centers[j].append(yc)
            ax.text(float(x[i]), yc, _fmt_num(v), ha="center", va="center", fontsize=9.0, color=txt_color, zorder=4)
        bottom = bottom + np.nan_to_num(y, nan=0.0)

    totals = bottom.copy()
    total_label_tops: list[float] = []
    for i, total in enumerate(totals):
        y_lbl = float(total) + max(abs(float(total)) * 0.03, 0.25)
        total_label_tops.append(y_lbl)
        ax.text(
            float(x[i]),
            y_lbl,
            _fmt_num(total),
            ha="center",
            va="bottom",
            fontsize=10.0,
            fontweight="bold" if i == n - 1 else "normal",
            color="#2f2f2f",
            zorder=5,
            clip_on=False,
        )

    # Brackets aligned on top for consecutive comparisons
    if n >= 2:
        abs_max = float(np.nanmax(np.abs(totals))) if np.isfinite(np.nanmax(np.abs(totals))) else 0.0
        offset_y = max(abs_max * 0.10, 0.5)
        bracket_h = max(abs_max * 0.03, 0.5)
        top_base = max(total_label_tops) + max(abs_max * 0.20, 1.0)
        max_text_y: float | None = None

        for i in range(1, n):
            prev = float(totals[i - 1])
            curr = float(totals[i])
            if not np.isfinite(prev) or not np.isfinite(curr) or prev == 0:
                continue
            pct = (curr / prev - 1.0) * 100.0
            lbl = f"{pct:+.1f}%".replace(".", ",")
            x1, x2 = float(x[i - 1]), float(x[i])
            y = top_base
            ax.plot([x1, x1, x2, x2], [y, y + bracket_h, y + bracket_h, y], color="#2f2f2f", linewidth=1.2, zorder=4)
            ty = y + bracket_h + offset_y * 0.25
            ax.text((x1 + x2) / 2.0, ty, lbl, ha="center", va="bottom", fontsize=9.0, color="#2f2f2f", zorder=5)
            max_text_y = ty if max_text_y is None else max(max_text_y, ty)

        if max_text_y is not None:
            ymin, ymax = ax.get_ylim()
            ax.set_ylim(ymin, max(ymax, max_text_y + offset_y * 1.2))

    # Inline legend on left
    x_leg = float(x.min()) - 0.9
    for j, name in enumerate(series_names):
        y_ref = centers[j][-1] if centers[j] else float("nan")
        if not np.isfinite(y_ref):
            for yc in centers[j]:
                if np.isfinite(yc):
                    y_ref = yc
                    break
        if not np.isfinite(y_ref):
            continue
        ax.scatter([x_leg], [float(y_ref)], s=90.0, marker="s", color=colors[j], edgecolors="none", zorder=6)
        ax.text(x_leg + 0.12, float(y_ref), str(name), ha="left", va="center", fontsize=9.0, color="#2f2f2f", zorder=6, clip_on=False)

    ax.set_xlim(float(x.min()) - 1.2, float(x.max()) + 0.65)
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


def _plot_media_percent(
    *,
    xlabels: list[str],
    values: list[float],  # fractions
    output_path: Path,
) -> None:
    import matplotlib.pyplot as plt

    vals = np.asarray(values, dtype=float) * 100.0
    x = np.arange(len(vals), dtype=float)

    fig, ax = plt.subplots(figsize=(10, 4.8), dpi=240)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    bars = ax.bar(x, vals, width=0.62, color="#123a7a", edgecolor="none", zorder=2)
    for i, (rect, v) in enumerate(zip(bars, vals)):
        ax.text(
            rect.get_x() + rect.get_width() / 2,
            rect.get_height(),
            _fmt_pct(v),
            ha="center",
            va="bottom",
            fontsize=10.0,
            fontweight="bold" if i == len(vals) - 1 else "normal",
            color="#2f2f2f",
            zorder=4,
            clip_on=False,
        )

    ax.set_xticks(x)
    ax.set_xticklabels(xlabels, fontsize=10.0)
    ax.tick_params(axis="x", bottom=False, pad=8)
    ax.set_yticks([])
    for s in ("left", "right", "top", "bottom"):
        ax.spines[s].set_visible(False)
    ax.grid(False)
    ax.margins(x=0.05, y=0.22)

    fig.tight_layout(pad=0.2)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=450, transparent=True, bbox_inches="tight", pad_inches=0.05)
    close_figure(fig)


def generate_slide13_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
    """Slide 13: originação (tri + 9M) e médias (tri + 9M)."""

    output_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(filename=xlsx_path, data_only=True)
    sheet_name = "slide_13"
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Aba não encontrada: {sheet_name!r}. Disponíveis: {wb.sheetnames}")
    ws = wb[sheet_name]

    labels = [("" if v is None else str(v)).strip() for v in _read_range_row(ws, "C5:G5")]
    s1 = to_float_list(_read_range_row(ws, "C6:G6"))
    s2 = to_float_list(_read_range_row(ws, "C7:G7"))
    values = np.column_stack((np.asarray(s1, dtype=float), np.asarray(s2, dtype=float)))
    series_names = [("" if v is None else str(v)).strip() for v in _read_range_row(ws, "B6:B7")]

    generated: list[Path] = []

    out26 = output_dir / "26_originacao_veiculos_trimestres.png"
    _plot_stacked_originacao(
        xlabels=labels[:3],
        series_names=series_names,
        values=values[:3, :],
        output_path=out26,
    )
    generated.append(out26)

    out27 = output_dir / "27_originacao_veiculos_9m.png"
    _plot_stacked_originacao(
        xlabels=labels[3:],
        series_names=series_names,
        values=values[3:, :],
        output_path=out27,
    )
    generated.append(out27)

    media_labels = [("" if v is None else str(v)).strip() for v in _read_range_row(ws, "K5:O5")]
    media_values = to_float_list(_read_range_row(ws, "K6:O6"))

    out28 = output_dir / "28_medias_trimestres.png"
    _plot_media_percent(
        xlabels=media_labels[:3],
        values=media_values[:3],
        output_path=out28,
    )
    generated.append(out28)

    out29 = output_dir / "29_medias_9m.png"
    _plot_media_percent(
        xlabels=media_labels[3:],
        values=media_values[3:],
        output_path=out29,
    )
    generated.append(out29)

    return generated


if __name__ == "__main__":
    xlsx = Path("testing.xlsx")
    out = Path(".")
    if xlsx.exists():
        files = generate_slide13_charts(xlsx_path=xlsx, output_dir=out)
        print(f"Gerados: {files}")
    else:
        print(f"Arquivo {xlsx} não encontrado")
