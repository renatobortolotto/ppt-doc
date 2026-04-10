import json
from pathlib import Path
from typing import Any, Dict

from genai_framework.decorators import file_input_route  # framework corporativo
from genai_framework.models import FileInput  # framework corporativo

from presentation_builder import build_presentation_from_bytes, load_job_config
from src.application.response_safety import build_error_response, serialize_build_response


def _repo_root() -> Path:
    return Path(__file__).resolve().parent


def _parse_llm_payload(llm_response_bytes: bytes) -> object:
    if not llm_response_bytes:
        raise ValueError("JSON da LLM vazio")
    try:
        return json.loads(llm_response_bytes.decode("utf-8"))
    except UnicodeDecodeError as exc:
        raise ValueError("JSON da LLM deve estar em UTF-8") from exc
    except json.JSONDecodeError as exc:
        raise ValueError("JSON da LLM invalido") from exc


def compose_presentation_from_inputs(xlsx_bytes: bytes, llm_response_bytes: bytes) -> Dict[str, Any]:
    repo_root = _repo_root()
    cfg = load_job_config(repo_root)
    llm_payload = _parse_llm_payload(llm_response_bytes)
    pptx_bytes, result = build_presentation_from_bytes(
        repo_root=repo_root,
        cfg=cfg,
        xlsx_bytes=xlsx_bytes,
        llm_payload=llm_payload,
    )
    filename = Path(str(cfg.get("api_output_filename") or cfg.get("pptx_output") or "presentation.updated.pptx")).name
    return serialize_build_response(
        pptx_bytes=pptx_bytes,
        filename=filename,
        result=result,
    )


def compose_presentation_files(xlsx_file: FileInput, llm_response_file: FileInput) -> Dict[str, Any]:
    try:
        return compose_presentation_from_inputs(xlsx_file.content, llm_response_file.content)
    except ValueError:
        return build_error_response(
            error="Falha ao montar o PowerPoint a partir do XLSX e do JSON da LLM.",
            error_code="invalid_request",
            details="Verifique o payload enviado e os arquivos recebidos.",
        )
    except Exception:
        return build_error_response(
            error="Falha ao montar o PowerPoint a partir do XLSX e do JSON da LLM.",
            error_code="build_failed",
            details="Consulte os logs do servidor para detalhes tecnicos.",
        )


@file_input_route("compose_presentation")
def compose_presentation(xlsx_file: FileInput, llm_response_file: FileInput):
    return compose_presentation_files(xlsx_file, llm_response_file)

        
