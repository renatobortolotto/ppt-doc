from __future__ import annotations

import base64
import re
from html import escape, unescape
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


ATTACHMENT_FILENAME_UNSAFE_CHARS_RE = re.compile(r"""[\x00-\x1f\x7f<>:"/\\|?*&'`;]""")
DEFAULT_ATTACHMENT_FILENAME = "presentation.updated.pptx"
MAX_ATTACHMENT_FILENAME_LENGTH = 180


def sanitize_response_text(value: Any) -> str:
    return escape(unescape(str(value)), quote=True)


def sanitize_response_optional_text(value: Any) -> str | None:
    if value is None:
        return None
    return sanitize_response_text(value)


def sanitize_response_value(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, Path):
        return sanitize_response_text(value)
    if isinstance(value, str):
        return sanitize_response_text(value)
    if isinstance(value, Mapping):
        return {str(key): sanitize_response_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [sanitize_response_value(item) for item in value]
    return sanitize_response_text(value)


def _safe_attachment_filename(filename: str | Path | None) -> str:
    raw_name = Path(str(filename or "")).name
    safe_name = ATTACHMENT_FILENAME_UNSAFE_CHARS_RE.sub("_", raw_name)
    safe_name = re.sub(r"\s+", " ", safe_name).strip(" ._")

    if not safe_name:
        safe_name = DEFAULT_ATTACHMENT_FILENAME

    if Path(safe_name).suffix.lower() != ".pptx":
        stem = safe_name.rstrip(".") or Path(DEFAULT_ATTACHMENT_FILENAME).stem
        safe_name = f"{stem}.pptx"

    if len(safe_name) > MAX_ATTACHMENT_FILENAME_LENGTH:
        suffix = Path(safe_name).suffix or ".pptx"
        stem_limit = MAX_ATTACHMENT_FILENAME_LENGTH - len(suffix)
        stem = Path(safe_name).stem[:stem_limit].rstrip(" ._")
        safe_name = f"{stem or Path(DEFAULT_ATTACHMENT_FILENAME).stem}{suffix}"

    return safe_name


def serialize_build_response(*, pptx_bytes: bytes, filename: str, result: Any) -> Dict[str, Any]:
    chart_failures = getattr(result, "chart_failures", ())
    text_field_failures = getattr(result, "text_field_failures", ())
    payload = {
        "filename": sanitize_response_text(filename),
        "contentType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "pptxBase64": base64.b64encode(pptx_bytes).decode("ascii"),
        "summary": {
            "outputPath": sanitize_response_text(result.output_path),
            "replacedPictures": result.replaced_pictures,
            "replacedPlaceholders": result.replaced_placeholders,
            "replacedText": result.replaced_text,
            "generatedChartCount": result.generated_chart_count,
            "chartFailureCount": len(chart_failures),
            "chartFailures": [
                {
                    "generatorKey": sanitize_response_text(failure.generator_key),
                    "label": sanitize_response_text(failure.label),
                    "outputFiles": [
                        sanitize_response_text(output_file)
                        for output_file in failure.output_files
                    ],
                    "error": sanitize_response_text(failure.error),
                }
                for failure in chart_failures
            ],
            "textFieldFailureCount": len(text_field_failures),
            "textFieldFailures": [
                {
                    "fieldId": sanitize_response_text(failure.field_id),
                    "sheet": sanitize_response_optional_text(failure.sheet),
                    "range": sanitize_response_text(failure.a1_range),
                    "error": sanitize_response_text(failure.error),
                }
                for failure in text_field_failures
            ],
            "appliedTextKeys": [
                sanitize_response_text(key)
                for key in getattr(result, "applied_text_keys", ())
            ],
        },
    }
    return payload


def build_error_response(*, error: str, error_code: str, details: str) -> Dict[str, Any]:
    return {
        "error": sanitize_response_text(error),
        "errorCode": sanitize_response_text(error_code),
        "details": sanitize_response_text(details),
    }


def build_file_response(*, body: bytes, filename: str, content_type: str) -> Response:
    safe_name = _safe_attachment_filename(filename)

    ascii_name = _safe_attachment_filename(
        safe_name.encode("ascii", "ignore").decode("ascii")
    )
    content_disposition = f'attachment; filename="{ascii_name}"'
    if safe_name != ascii_name:
        content_disposition += f"; filename*=UTF-8''{quote(safe_name, safe='')}"

    headers = {
        "Content-Disposition": content_disposition,
        "Content-Length": str(len(body)),
    }
    return Response(response=body, status=200, headers=headers, mimetype=content_type)
