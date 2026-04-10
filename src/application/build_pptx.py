from __future__ import annotations

import base64
import binascii
import json
import logging
from pathlib import Path
from typing import Any, Dict, Mapping

from presentation_builder import build_presentation_from_bytes, load_job_config
from src.infrastructure.framework_compat import Resource, request


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


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
    chart_failures = getattr(result, "chart_failures", ())
    text_field_failures = getattr(result, "text_field_failures", ())
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
            "chartFailureCount": len(chart_failures),
            "chartFailures": [
                {
                    "generatorKey": failure.generator_key,
                    "label": failure.label,
                    "outputFiles": list(failure.output_files),
                    "error": failure.error,
                }
                for failure in chart_failures
            ],
            "textFieldFailureCount": len(text_field_failures),
            "textFieldFailures": [
                {
                    "fieldId": failure.field_id,
                    "sheet": failure.sheet,
                    "range": failure.a1_range,
                    "error": failure.error,
                }
                for failure in text_field_failures
            ],
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
    filename = Path(
        str(cfg.get("api_output_filename") or cfg.get("pptx_output") or "presentation.updated.pptx")
    ).name
    return _serialize_build_response(
        pptx_bytes=pptx_bytes,
        filename=filename,
        result=result,
    )


def _read_upload_bytes(upload: Any) -> bytes:
    if upload is None:
        return b""
    if isinstance(upload, (bytes, bytearray)):
        return bytes(upload)
    if hasattr(upload, "content"):
        content = upload.content
        if isinstance(content, (bytes, bytearray)):
            return bytes(content)
        raise TypeError("Campo de arquivo com atributo content invalido")
    if hasattr(upload, "read"):
        data = upload.read()
        try:
            upload.seek(0)
        except Exception:
            pass
        if isinstance(data, (bytes, bytearray)):
            return bytes(data)
        raise TypeError("Campo de upload retornou conteudo nao binario")
    raise TypeError("Campo de upload nao suportado")


def _first_value(mapping: Mapping[str, Any] | None, *keys: str) -> Any:
    if not mapping:
        return None
    for key in keys:
        if key in mapping and mapping[key] not in (None, ""):
            return mapping[key]
    return None


def _decode_base64_field(value: Any, *, field_name: str) -> bytes:
    if value in (None, ""):
        return b""
    if not isinstance(value, str):
        raise ValueError(f"Campo {field_name} deve ser string em base64")
    try:
        return base64.b64decode(value, validate=True)
    except (ValueError, binascii.Error) as exc:
        raise ValueError(f"Campo {field_name} nao contem base64 valido") from exc


def _json_to_bytes(value: Any, *, field_name: str) -> bytes:
    if value in (None, ""):
        return b""
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    if isinstance(value, str):
        return value.encode("utf-8")
    try:
        return json.dumps(value, ensure_ascii=False).encode("utf-8")
    except TypeError as exc:
        raise ValueError(f"Campo {field_name} nao pode ser serializado em JSON") from exc


def _extract_request_payload(request_obj: Any) -> tuple[bytes, bytes]:
    files = getattr(request_obj, "files", None) or {}
    form = getattr(request_obj, "form", None) or {}

    xlsx_upload = _first_value(files, "xlsx_file", "xlsx")
    llm_upload = _first_value(files, "llm_response_file", "llm_json_file", "llm_response")
    llm_form_value = _first_value(form, "llm_response_json", "llm_response")

    if xlsx_upload is not None:
        xlsx_bytes = _read_upload_bytes(xlsx_upload)
        llm_bytes = _read_upload_bytes(llm_upload) if llm_upload is not None else _json_to_bytes(
            llm_form_value,
            field_name="llm_response_json",
        )
        if not llm_bytes:
            raise ValueError(
                "Envie o JSON da LLM em llm_response_file ou llm_response_json junto com o xlsx_file"
            )
        return xlsx_bytes, llm_bytes

    json_payload = None
    if hasattr(request_obj, "get_json"):
        try:
            json_payload = request_obj.get_json(silent=True)
        except TypeError:
            json_payload = request_obj.get_json()

    if isinstance(json_payload, dict):
        xlsx_base64 = _first_value(
            json_payload,
            "xlsxBase64",
            "xlsx_base64",
            "xlsx_file_base64",
            "file",
        )
        llm_base64 = _first_value(
            json_payload,
            "llmResponseBase64",
            "llm_response_base64",
            "llm_response_file_base64",
        )
        llm_inline = _first_value(
            json_payload,
            "llmResponse",
            "llm_response",
            "data",
        )
        xlsx_bytes = _decode_base64_field(xlsx_base64, field_name="xlsxBase64")
        if llm_base64:
            llm_bytes = _decode_base64_field(llm_base64, field_name="llmResponseBase64")
        else:
            llm_bytes = _json_to_bytes(llm_inline, field_name="llmResponse")
        if not xlsx_bytes or not llm_bytes:
            raise ValueError(
                "No corpo JSON, envie xlsxBase64 + llmResponseBase64 ou xlsxBase64 + llmResponse"
            )
        return xlsx_bytes, llm_bytes

    raise ValueError(
        "Requisicao sem payload suportado. Use multipart/form-data com xlsx_file + llm_response_file "
        "ou application/json com xlsxBase64 + llmResponseBase64"
    )


def handle_build_pptx_request(request_obj: Any | None = None) -> tuple[Dict[str, Any], int]:
    active_request = request_obj or request
    try:
        xlsx_bytes, llm_response_bytes = _extract_request_payload(active_request)
        response = compose_presentation_from_inputs(xlsx_bytes, llm_response_bytes)
        return response, 200
    except ValueError as exc:
        return {
            "error": "Requisicao invalida para build-pptx.",
            "details": str(exc),
        }, 400
    except Exception as exc:
        logging.exception("Falha ao processar a requisicao build-pptx")
        return {
            "error": "Falha ao montar o PowerPoint a partir do XLSX e do JSON da LLM.",
            "details": str(exc),
        }, 500


class BuildPptxResource(Resource):
    def post(self):
        return handle_build_pptx_request()
