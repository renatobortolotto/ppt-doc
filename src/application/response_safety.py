from __future__ import annotations

import base64
from html import escape
from pathlib import Path
from typing import Any, Dict, Mapping
from urllib.parse import quote

try:
    from flask import Response
except ModuleNotFoundError:
    class Response:
        def __init__(self, response=b"", status=200, headers=None, mimetype=None):
            self.status_code = status
            self.headers = dict(headers or {})
            self.mimetype = mimetype
            if mimetype and "Content-Type" not in self.headers:
                self.headers["Content-Type"] = mimetype
            self.data = response if isinstance(response, (bytes, bytearray)) else str(response).encode("utf-8")

        def get_data(self, as_text: bool = False):
            if as_text:
                return self.data.decode("utf-8")
            return self.data


def sanitize_response_value(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, Path):
        return escape(str(value), quote=True)
    if isinstance(value, str):
        return escape(value, quote=True)
    if isinstance(value, Mapping):
        return {str(key): sanitize_response_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [sanitize_response_value(item) for item in value]
    return escape(str(value), quote=True)


def serialize_build_response(*, pptx_bytes: bytes, filename: str, result: Any) -> Dict[str, Any]:
    chart_failures = getattr(result, "chart_failures", ())
    text_field_failures = getattr(result, "text_field_failures", ())
    payload = {
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
            "appliedTextKeys": list(getattr(result, "applied_text_keys", ())),
        },
    }
    return sanitize_response_value(payload)


def build_error_response(*, error: str, error_code: str, details: str) -> Dict[str, Any]:
    return sanitize_response_value(
        {
            "error": error,
            "errorCode": error_code,
            "details": details,
        }
    )


def build_file_response(*, body: bytes, filename: str, content_type: str) -> Response:
    safe_name = Path(str(filename)).name.replace('"', "")
    if not safe_name:
        safe_name = "presentation.updated.pptx"

    ascii_name = safe_name.encode("ascii", "ignore").decode("ascii") or "presentation.updated.pptx"
    content_disposition = f'attachment; filename="{ascii_name}"'
    if safe_name != ascii_name:
        content_disposition += f"; filename*=UTF-8''{quote(safe_name)}"

    headers = {
        "Content-Disposition": content_disposition,
        "Content-Length": str(len(body)),
    }
    return Response(response=body, status=200, headers=headers, mimetype=content_type)
