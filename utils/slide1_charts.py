"""Compatibilidade legada para imports antigos.

O modulo original foi removido do fluxo atual.
"""

from __future__ import annotations

from pathlib import Path


def generate_slide1_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
    raise RuntimeError(
        "O grafico legado do slide 1/7 foi removido do workspace atual e nao pode mais ser gerado"
    )
