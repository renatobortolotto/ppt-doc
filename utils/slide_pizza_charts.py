from __future__ import annotations

from pathlib import Path

from utils.charts_common import (
    ExcelBarChartSpec,
    ExcelDonutChartSpec,
    close_figure,
    plot_bar_from_excel,
    plot_donut_from_excel,
)


def generate_pizza_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
    """Gera gráficos da worksheet 'Pizza Teste': donut + barras (10, 11, 12)."""

    output_dir.mkdir(parents=True, exist_ok=True)

    generated: list[Path] = []

    # Cores para os segmentos internos do donut
    inner_colors = [
        "#1f3a8a",  # Veículos Leves Usados (azul escuro)
        "#b11226",  # Corporate (vermelho)
        "#d64550",  # Large e Instituições (vermelho claro)
        "#b0dfe5",  # Demais (azul bem claro)
        "#7ec8e3",  # PME (azul claro)
        "#3cb4ac",  # EGV (verde água)
        "#2ca6a4",  # Cartões (verde água escuro)
        "#6bd4c6",  # Painéis Solares (verde claro)
        "#4f9da6",  # Motos, Pesados e Novos (azul esverdeado)
    ]

    # Cores para as categorias externas do donut
    outer_colors = [
        "#1f3a8a",  # Veículos Leves - azul escuro
        "#b11226",  # Atacado - vermelho
        "#3cb4ac",  # Growth - verde água
    ]

    # 10) Gráfico de Carteira de Crédito Ampliada (donut)
    fig, _ax = plot_donut_from_excel(
        ExcelDonutChartSpec(
            file_path=xlsx_path,
            sheet_name="Pizza Teste",
            categories_range="A2:A10",
            labels_range="B2:B10",
            values_range="C2:C10",
            center_text="Carteira\nAmpliada\nR$ 92.7 bi",
            title=None,
            inner_colors=inner_colors,
            outer_colors=outer_colors,
            output_path=output_dir / "10_pizza_carteira.png",
            figsize=(16, 12),
            font_scale=1.5,
        )
    )
    close_figure(fig)
    generated.append(output_dir / "10_pizza_carteira.png")

    # 11) Barras - Trimestres (H3:J3 labels, H4:J4 valores)
    fig, _ax = plot_bar_from_excel(
        ExcelBarChartSpec(
            file_path=xlsx_path,
            sheet_name="Pizza Teste",
            values_range="H4:J4",
            xlabels_range="H3:J3",
            title=None,
            highlight_last=True,
            bar_color="#123a7a",
            show_delta_pct=True,
            show_delta_bracket=True,
            delta_pairs=((0, 1), (1, 2)),
            font_scale=1.5,
            output_path=output_dir / "11_pizza_trimestres.png",
        )
    )
    close_figure(fig)
    generated.append(output_dir / "11_pizza_trimestres.png")

    # 12) Barras - 9M (K3:L3 labels, K4:L4 valores)
    fig, _ax = plot_bar_from_excel(
        ExcelBarChartSpec(
            file_path=xlsx_path,
            sheet_name="Pizza Teste",
            values_range="K4:L4",
            xlabels_range="K3:L3",
            title=None,
            highlight_last=True,
            bar_color="#123a7a",
            show_delta_pct=True,
            show_delta_bracket=True,
            fixed_slot_count=9,
            font_scale=1.5,
            output_path=output_dir / "12_pizza_9m.png",
        )
    )
    close_figure(fig)
    generated.append(output_dir / "12_pizza_9m.png")

    return generated


if __name__ == "__main__":
    # Teste local
    xlsx = Path("testing.xlsx")
    out = Path(".")
    if xlsx.exists():
        files = generate_pizza_charts(xlsx_path=xlsx, output_dir=out)
        print(f"Gerados: {files}")
    else:
        print(f"Arquivo {xlsx} não encontrado")
