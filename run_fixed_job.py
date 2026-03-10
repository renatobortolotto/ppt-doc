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


def _resolve_path(repo_root: Path, p: str) -> Path:
    path = Path(p).expanduser()
    if path.is_absolute():
        return path
    return (repo_root / path).resolve()


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


def _call_analyze_api(*, api_url: str, xlsx_path: Path, cfg: Dict[str, Any]) -> object:
    """Call the deployed FastAPI endpoint with the XLSX as multipart.

    Returns the parsed JSON (already decoded).
    """

    try:
        import requests  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise RuntimeError(
            "Dependência 'requests' não instalada. Instale via requirements.txt."
        ) from exc

    field = str(cfg.get("api_file_field") or "file")
    timeout = float(cfg.get("api_timeout_seconds") or 180)
    headers = cfg.get("api_headers")
    if headers is None:
        headers = {}
    if not isinstance(headers, dict):
        raise ValueError("api_headers deve ser um objeto JSON")

    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    with xlsx_path.open("rb") as f:
        files = {
            field: (
                xlsx_path.name,
                f,
                content_type,
            )
        }
        logging.info("Chamando API: %s (field=%s)", api_url, field)
        resp = requests.post(api_url, files=files, headers=headers, timeout=timeout)

    if resp.status_code >= 400:
        raise RuntimeError(
            f"API retornou {resp.status_code}: {resp.text[:2000]}"
        )

    try:
        return resp.json()
    except Exception as exc:
        raise RuntimeError(
            f"API não retornou JSON válido. Body (parcial): {resp.text[:2000]}"
        ) from exc


def _maybe_fetch_llm_response(repo_root: Path, cfg: Dict[str, Any], xlsx_path: Path) -> None:
    """If api_url is configured, call the API and persist JSON to llm_response_json."""

    api_url = cfg.get("api_url")
    if not api_url:
        return

    api_url_s = str(api_url)
    payload = _call_analyze_api(api_url=api_url_s, xlsx_path=xlsx_path, cfg=cfg)

    out_path = cfg.get("llm_response_json")
    if not out_path:
        raise ValueError("Para usar api_url, configure também llm_response_json")

    out = _resolve_path(repo_root, str(out_path))
    out.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    logging.info("Resposta da API salva em: %s", str(out))


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
            "Quando informado, ele sobrescreve o llm_response_json configurado no job."
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
    else:
        # If configured, fetch the LLM JSON from FastAPI automatically.
        _maybe_fetch_llm_response(repo_root, cfg, xlsx_path)

    result = build_presentation(
        repo_root=repo_root,
        cfg=cfg,
        xlsx_path=xlsx_path,
        llm_payload=llm_payload,
        skip_charts=bool(args.skip_charts),
    )

    logging.info(
        "OK: gerado %s (pictures=%d text=%d charts=%d)",
        str(result.output_path),
        result.replaced_pictures,
        result.replaced_text,
        result.generated_chart_count,
    )


if __name__ == "__main__":
    main()
