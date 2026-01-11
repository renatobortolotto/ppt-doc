from __future__ import annotations

from pathlib import Path

from utils.charts_common import ExcelBarChartSpec, close_figure, plot_bar_from_excel, plot_line_from_excel


def generate_slide1_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
    """Slide 1: gera 01, 02, 03, 04."""

    output_dir.mkdir(parents=True, exist_ok=True)

    generated: list[Path] = []

    # 01) Lucro líquido - Trimestres
    fig, _ax = plot_bar_from_excel(
        ExcelBarChartSpec(
            file_path=xlsx_path,
            sheet_name="DRE Saida",
            values_range="C18:K18",
            xlabels_range="C3:K3",
            ylabel_cell="B18",
            title=None,
            highlight_last=True,
            show_delta_pct=True,
            show_delta_bracket=True,
            delta_pairs=((-2, -1), (-5, -2)),
            output_path=output_dir / "01_lucro_trimestres.png",
        )
    )
    close_figure(fig)
    generated.append(output_dir / "01_lucro_trimestres.png")

    # 02) Lucro líquido - 9M
    fig, _ax = plot_bar_from_excel(
        ExcelBarChartSpec(
            file_path=xlsx_path,
            sheet_name="DRE Saida",
            values_range="L18:M18",
            xlabels_range="L3:M3",
            ylabel_cell="C18",
            title=None,
            highlight_last=True,
            show_delta_pct=True,
            show_delta_bracket=True,
            fixed_slot_count=9,
            output_path=output_dir / "02_lucro_9m.png",
        )
    )
    close_figure(fig)
    generated.append(output_dir / "02_lucro_9m.png")

    # 03) ROE - Trimestres
    fig, _ax = plot_line_from_excel(
        file_path=xlsx_path,
        sheet_name="DRE Saida",
        values_range="C20:K20",
        xlabels_range="C3:K3",
        output_path=output_dir / "03_roe_trimestres.png",
        fmt_as_percent=True,
        y_baseline=0.0,
        y_expand=0.10,
        smooth=True,
    )
    close_figure(fig)
    generated.append(output_dir / "03_roe_trimestres.png")

    # 04) ROE - 9M
    fig, _ax = plot_line_from_excel(
        file_path=xlsx_path,
        sheet_name="DRE Saida",
        values_range="L20:M20",
        xlabels_range="L3:M3",
        output_path=output_dir / "04_roe_9m.png",
        fmt_as_percent=True,
        smooth=True,
        line_width=5.0,
        label_fontsize=24.0,
        marker_size=96.0,
        label_offset_pts=18.0,
    )
    close_figure(fig)
    generated.append(output_dir / "04_roe_9m.png")

    return generated
