import base64
import json
from pathlib import Path
from typing import Any, Dict

from genai_framework.decorators import file_input_route  # framework corporativo
from genai_framework.models import FileInput  # framework corporativo

from presentation_builder import build_presentation_from_bytes, load_job_config


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


def _serialize_build_response(*, pptx_bytes: bytes, filename: str, result) -> Dict[str, Any]:
    return {
        "filename": filename,
        "contentType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "pptxBase64": base64.b64encode(pptx_bytes).decode("ascii"),
        "summary": {
            "outputPath": str(result.output_path),
            "replacedPictures": result.replaced_pictures,
            "replacedPlaceholders": result.replaced_placeholders,
            "replacedText": result.replaced_text,
            "generatedChartCount": result.generated_chart_count,
            "appliedTextKeys": list(result.applied_text_keys),
        },
    }


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
    return _serialize_build_response(
        pptx_bytes=pptx_bytes,
        filename=filename,
        result=result,
    )


def compose_presentation_files(xlsx_file: FileInput, llm_response_file: FileInput) -> Dict[str, Any]:
    try:
        return compose_presentation_from_inputs(xlsx_file.content, llm_response_file.content)
    except Exception as exc:
        return {
            "error": "Falha ao montar o PowerPoint a partir do XLSX e do JSON da LLM.",
            "details": str(exc),
        }


@file_input_route("compose_presentation")
def compose_presentation(xlsx_file: FileInput, llm_response_file: FileInput):
    return compose_presentation_files(xlsx_file, llm_response_file)

        
