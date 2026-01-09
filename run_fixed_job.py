from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path
from typing import Any, Dict

from update_ppt import _flatten_text_payload, update_presentation
from utils.xlsx_text_fields import extract_xlsx_to_text_mapping, parse_text_fields_json


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


def _load_llm_mapping(repo_root: Path, cfg: Dict[str, Any]) -> Dict[str, str]:
    llm_path = cfg.get("llm_response_json")
    if not llm_path:
        return {}

    path = _resolve_path(repo_root, str(llm_path))
    if not path.exists():
        raise FileNotFoundError(f"LLM response JSON não encontrado: {path}")

    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, dict) and "response" in payload and isinstance(payload["response"], dict):
        payload = payload["response"]
    return _flatten_text_payload(payload)


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
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(
        description=(
            "Job fixo: recebe apenas o XLSX e atualiza o PPT usando configs em config/*.json.\n\n"
            "Você edita config/job_config.json e config/text_fields.json uma única vez."
        )
    )
    parser.add_argument("--xlsx", required=True, help="Caminho do XLSX de entrada")
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parent
    cfg = _load_job_config(repo_root)

    xlsx_path = Path(args.xlsx).expanduser().resolve()
    if not xlsx_path.exists():
        raise FileNotFoundError(f"XLSX não encontrado: {xlsx_path}")

    # If configured, fetch the LLM JSON from FastAPI automatically.
    _maybe_fetch_llm_response(repo_root, cfg, xlsx_path)

    pptx_template = _resolve_path(repo_root, str(cfg.get("pptx_template")))
    pptx_output = _resolve_path(repo_root, str(cfg.get("pptx_output")))
    images_dir = _resolve_path(repo_root, str(cfg.get("images_dir", ".")))
    allow_placeholder_text = bool(cfg.get("allow_placeholder_text", False))

    text_fields_config = _resolve_path(repo_root, str(cfg.get("text_fields_config", "config/text_fields.json")))

    # XLSX-derived fields
    default_sheet_from_config, specs = parse_text_fields_json(text_fields_config)
    text_mapping = extract_xlsx_to_text_mapping(
        xlsx_path,
        specs,
        default_sheet=default_sheet_from_config,
    )

    # LLM-derived fields (optional): merge selected keys only
    raw_text_cfg = json.loads(text_fields_config.read_text(encoding="utf-8"))
    llm_fields: list[str] = []
    if isinstance(raw_text_cfg, dict):
        lf = raw_text_cfg.get("llm_fields") or raw_text_cfg.get("from_llm")
        if isinstance(lf, list):
            llm_fields = [str(x) for x in lf]

    llm_mapping = _load_llm_mapping(repo_root, cfg)
    if llm_fields:
        allowed = set(llm_fields)
        llm_mapping = {k: v for k, v in llm_mapping.items() if k in allowed}

    # LLM overrides XLSX for same keys
    text_mapping.update(llm_mapping)

    (
        replaced_pictures,
        replaced_placeholders,
        replaced_text,
        _replaced_files,
        _missing_files,
        _applied_text_keys,
    ) = update_presentation(
        pptx_path=pptx_template,
        output_path=pptx_output,
        images_dir=images_dir,
        allow_placeholder_text=allow_placeholder_text,
        text_json=None,
        text_payload=text_mapping,
    )

    logging.info(
        "OK: gerado %s (pictures=%d text=%d)",
        str(pptx_output),
        replaced_pictures,
        replaced_text,
    )


if __name__ == "__main__":
    main()
