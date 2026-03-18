from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any, Dict

from presentation_builder import build_presentation


def _add_handler_if_missing(logger: logging.Logger, handler: logging.Handler) -> None:
    handler_type = type(handler)
    for existing in logger.handlers:
        if isinstance(existing, handler_type):
            return
    logger.addHandler(handler)


def _configure_logging(level: str, *, log_file: str | None = None) -> None:
    """Configure logging reliably even if the host already configured handlers.

    In some corporate runtimes, `logging.basicConfig()` becomes a no-op because a handler
    already exists. This function ensures we always have a stream handler and level set.
    """

    numeric = getattr(logging, level.upper(), None)
    if not isinstance(numeric, int):
        numeric = logging.INFO

    root = logging.getLogger()
    fmt = logging.Formatter("%(levelname)s: %(message)s")

    # Python 3.8+: use force=True to override hostile pre-configured handlers.
    try:
        handlers: list[logging.Handler] = [logging.StreamHandler(stream=sys.stderr)]
        handlers[0].setFormatter(fmt)

        if log_file:
            log_path = Path(log_file).expanduser().resolve()
            log_path.parent.mkdir(parents=True, exist_ok=True)
            file_handler = logging.FileHandler(log_path, encoding="utf-8")
            file_handler.setFormatter(fmt)
            handlers.append(file_handler)

        logging.basicConfig(level=numeric, handlers=handlers, force=True)
    except TypeError:
        # Older Python: no force= parameter.
        root.setLevel(numeric)

        stream_handler = logging.StreamHandler(stream=sys.stderr)
        stream_handler.setFormatter(fmt)
        _add_handler_if_missing(root, stream_handler)

        if log_file:
            log_path = Path(log_file).expanduser().resolve()
            log_path.parent.mkdir(parents=True, exist_ok=True)
            file_handler = logging.FileHandler(log_path, encoding="utf-8")
            file_handler.setFormatter(fmt)
            _add_handler_if_missing(root, file_handler)

    # Third-party libraries can be extremely noisy at DEBUG.
    logging.getLogger("matplotlib").setLevel(logging.WARNING)
    logging.getLogger("PIL").setLevel(logging.WARNING)


def _load_job_config(repo_root: Path) -> Dict[str, Any]:
    cfg_path = repo_root / "config" / "job_config.json"
    if not cfg_path.exists():
        raise FileNotFoundError(
            f"Config não encontrada: {cfg_path}. Edite uma vez e rode novamente."
        )
    raw = json.loads(cfg_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("job_config.json deve ser um objeto")
    return raw


def _parse_only_slides(raw_value: str) -> tuple[int, ...]:
    slides: list[int] = []

    for chunk in str(raw_value).split(","):
        part = chunk.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = [value.strip() for value in part.split("-", 1)]
            if not start_text.isdigit() or not end_text.isdigit():
                raise ValueError(
                    "only-slides deve usar numeros separados por virgula ou intervalos como 1-8"
                )
            start = int(start_text)
            end = int(end_text)
            if start <= 0 or end <= 0 or end < start:
                raise ValueError("only-slides precisa usar numeros positivos em ordem crescente")
            slides.extend(range(start, end + 1))
            continue

        if not part.isdigit():
            raise ValueError(
                "only-slides deve usar numeros separados por virgula ou intervalos como 1-8"
            )
        slide = int(part)
        if slide <= 0:
            raise ValueError("only-slides precisa usar numeros positivos")
        slides.append(slide)

    normalized = tuple(dict.fromkeys(slides))
    if not normalized:
        raise ValueError("only-slides nao pode ficar vazio")
    return normalized


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Job fixo: recebe o XLSX e opcionalmente o JSON da LLM, atualizando o PPT usando configs em config/*.json.\n\n"
            "Você edita config/job_config.json e config/text_fields.json uma única vez."
        )
    )
    parser.add_argument("--xlsx", required=True, help="Caminho do XLSX de entrada")
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Nível de log (DEBUG, INFO, WARNING, ERROR). Default: INFO",
    )
    parser.add_argument(
        "--log-file",
        default=None,
        help=(
            "Opcional: caminho para gravar logs em arquivo (útil em ambientes corporativos que escondem stdout/stderr). "
            "Também pode ser definido via env PPTDOC_LOG_FILE."
        ),
    )
    parser.add_argument(
        "--skip-charts",
        action="store_true",
        help="Não gera os PNGs antes de atualizar o PPT.",
    )
    parser.add_argument(
        "--llm-json",
        default=None,
        help=(
            "Opcional: caminho do JSON já gerado pela LLM. "
            "Quando omitido, o runner usa o llm_response_json configurado no job."
        ),
    )
    parser.add_argument(
        "--only-slides",
        default=None,
        help=(
            "Opcional: limita a geracao de graficos aos slides informados. "
            "Aceita lista e intervalos, por exemplo: 3,4,7,8 ou 1-8."
        ),
    )
    args = parser.parse_args()

    log_file = args.log_file or os.environ.get("PPTDOC_LOG_FILE")
    _configure_logging(str(args.log_level), log_file=log_file)
    logging.info("Logging inicializado (level=%s)%s", str(args.log_level).upper(), f" file={log_file}" if log_file else "")

    repo_root = Path(__file__).resolve().parent
    cfg = _load_job_config(repo_root)

    xlsx_path = Path(args.xlsx).expanduser().resolve()
    if not xlsx_path.exists():
        raise FileNotFoundError(f"XLSX não encontrado: {xlsx_path}")

    llm_payload = None
    if args.llm_json:
        llm_json_path = Path(args.llm_json).expanduser().resolve()
        if not llm_json_path.exists():
            raise FileNotFoundError(f"JSON da LLM não encontrado: {llm_json_path}")
        llm_payload = json.loads(llm_json_path.read_text(encoding="utf-8"))

    only_slides = None
    if args.only_slides:
        try:
            only_slides = _parse_only_slides(str(args.only_slides))
        except ValueError as exc:
            parser.error(str(exc))
        logging.info("Limitando geracao de graficos aos slides: %s", ", ".join(str(slide) for slide in only_slides))

    result = build_presentation(
        repo_root=repo_root,
        cfg=cfg,
        xlsx_path=xlsx_path,
        llm_payload=llm_payload,
        skip_charts=bool(args.skip_charts),
        only_slides=only_slides,
    )

    logging.info(
        "OK: gerado %s (pictures=%d text=%d charts=%d chart_failures=%d text_field_failures=%d)",
        str(result.output_path),
        result.replaced_pictures,
        result.replaced_text,
        result.generated_chart_count,
        len(result.chart_failures),
        len(result.text_field_failures),
    )
    if result.chart_failures:
        for failure in result.chart_failures:
            logging.warning(
                "Grafico com falha em %s: arquivos=%s erro=%s",
                failure.label,
                ", ".join(failure.output_files),
                failure.error,
            )
    if result.text_field_failures:
        for failure in result.text_field_failures:
            logging.warning(
                "Campo de texto com falha: field=%s sheet=%s range=%s erro=%s",
                failure.field_id,
                failure.sheet or "<sem-sheet>",
                failure.a1_range,
                failure.error,
            )


if __name__ == "__main__":
    main()
