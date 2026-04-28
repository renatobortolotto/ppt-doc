from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from presentation_builder import generate_chart_assets
from run_fixed_job import _configure_logging, _load_job_config, _parse_only_slides


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Gera somente os PNGs dos graficos usando o mesmo pipeline central do projeto, "
            "sem atualizar o PPT nem chamar o endpoint HTTP."
        )
    )
    parser.add_argument("--xlsx", required=True, help="Caminho do XLSX de entrada.")
    parser.add_argument(
        "--output-dir",
        default=None,
        help=(
            "Diretorio de saida dos PNGs. Quando omitido, usa a raiz do projeto."
        ),
    )
    parser.add_argument(
        "--only-slides",
        default=None,
        help=(
            "Opcional: limita a geracao aos slides informados. "
            "Aceita lista e intervalos, por exemplo: 4,8,11,12 ou 8-16."
        ),
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Retorna exit code 1 se algum gerador falhar.",
    )
    parser.add_argument(
        "--summary-json",
        default=None,
        help="Opcional: caminho para gravar um resumo JSON da execucao.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Nivel de log (DEBUG, INFO, WARNING, ERROR). Default: INFO",
    )
    parser.add_argument(
        "--log-file",
        default=None,
        help=(
            "Opcional: caminho para gravar logs em arquivo. "
            "Tambem pode ser definido via env PPTDOC_LOG_FILE."
        ),
    )
    return parser.parse_args()


def _resolve_output_dir(repo_root: Path, _cfg: dict[str, object], raw_output_dir: str | None) -> Path:
    if raw_output_dir:
        return Path(raw_output_dir).expanduser().resolve()
    return repo_root.resolve()


def _build_summary(*, xlsx_path: Path, output_dir: Path, result) -> dict[str, object]:
    return {
        "xlsx": str(xlsx_path),
        "outputDir": str(output_dir),
        "generatedCount": len(result.generated_files),
        "generatedFiles": [str(path) for path in result.generated_files],
        "failureCount": len(result.failures),
        "failures": [
            {
                "generatorKey": failure.generator_key,
                "label": failure.label,
                "outputFiles": list(failure.output_files),
                "error": failure.error,
            }
            for failure in result.failures
        ],
    }


def main() -> int:
    args = parse_args()

    log_file = args.log_file or os.environ.get("PPTDOC_LOG_FILE")
    _configure_logging(str(args.log_level), log_file=log_file)

    cfg = _load_job_config(REPO_ROOT)

    xlsx_path = Path(args.xlsx).expanduser().resolve()
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"XLSX nao encontrado: {xlsx_path}")

    output_dir = _resolve_output_dir(REPO_ROOT, cfg, args.output_dir)

    only_slides = None
    if args.only_slides:
        try:
            only_slides = _parse_only_slides(str(args.only_slides))
        except ValueError as exc:
            raise SystemExit(str(exc)) from exc
        logging.info(
            "Limitando geracao de graficos aos slides: %s",
            ", ".join(str(slide) for slide in only_slides),
        )

    logging.info("Gerando graficos a partir de %s", xlsx_path)
    logging.info("Diretorio de saida dos PNGs: %s", output_dir)

    result = generate_chart_assets(
        xlsx_path=xlsx_path,
        images_dir=output_dir,
        only_slides=only_slides,
    )

    summary = _build_summary(xlsx_path=xlsx_path, output_dir=output_dir, result=result)

    if args.summary_json:
        summary_path = Path(args.summary_json).expanduser().resolve()
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        logging.info("Resumo JSON salvo em %s", summary_path)

    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if result.failures:
        logging.warning(
            "Geracao concluida com falhas parciais: generated=%d failures=%d",
            len(result.generated_files),
            len(result.failures),
        )
        if args.strict:
            return 1
    else:
        logging.info("Geracao concluida sem falhas: generated=%d", len(result.generated_files))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
