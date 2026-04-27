import types
import unittest
from pathlib import Path
from unittest.mock import patch

import src.application.build_pptx as build_pptx


class _ContentUpload:
    def __init__(self, content, **attrs):
        self.content = content
        for key, value in attrs.items():
            setattr(self, key, value)


class _ReadUpload:
    def __init__(self, data, *, seek_raises=False):
        self._data = data
        self.seek_raises = seek_raises
        self.seek_calls = []

    def read(self):
        return self._data

    def seek(self, offset):
        self.seek_calls.append(offset)
        if self.seek_raises:
            raise RuntimeError("seek failed")


class _FilenameUpload:
    def __init__(self, **attrs):
        for key, value in attrs.items():
            setattr(self, key, value)


class _FakeRequest:
    def __init__(
        self,
        *,
        files=None,
        form=None,
        json_payload=None,
        type_error_on_silent=False,
    ):
        self.files = files or {}
        self.form = form or {}
        self.json_payload = json_payload
        self.type_error_on_silent = type_error_on_silent

    def get_json(self, silent=False):
        if silent and self.type_error_on_silent:
            raise TypeError("silent not supported")
        return self.json_payload


class TestBuildPptxApplication(unittest.TestCase):
    def _fake_build_result(self, name="saida.pptx"):
        return types.SimpleNamespace(
            output_path=Path(name),
            replaced_pictures=1,
            replaced_placeholders=2,
            replaced_text=3,
            generated_chart_count=4,
            chart_failures=(),
            text_field_failures=(),
            applied_text_keys=("slide1_title",),
        )

    def test_repo_root_points_to_project_root(self):
        self.assertEqual(
            build_pptx._repo_root(),
            Path(__file__).resolve().parents[1],
        )

    def test_parse_llm_payload_accepts_valid_json(self):
        self.assertEqual(
            build_pptx._parse_llm_payload(b'{"response":{"ok":true}}'),
            {"response": {"ok": True}},
        )

    def test_parse_llm_payload_rejects_empty_invalid_utf8_and_invalid_json(self):
        with self.assertRaisesRegex(ValueError, "vazio"):
            build_pptx._parse_llm_payload(b"")
        with self.assertRaisesRegex(ValueError, "UTF-8"):
            build_pptx._parse_llm_payload(b"\xff")
        with self.assertRaisesRegex(ValueError, "invalido"):
            build_pptx._parse_llm_payload(b"{")

    def test_build_presentation_artifact_loads_config_and_passes_inputs(self):
        fake_result = self._fake_build_result("relatorio.pptx")

        with (
            patch.object(build_pptx, "_repo_root", return_value=Path("/repo")) as root_mock,
            patch.object(build_pptx, "load_job_config", return_value={"cfg": True}) as cfg_mock,
            patch.object(
                build_pptx,
                "build_presentation_from_bytes",
                return_value=(b"pptx-bytes", fake_result),
            ) as build_mock,
        ):
            pptx_bytes, filename, result = build_pptx._build_presentation_artifact(
                b"xlsx-bytes",
                b'{"response":{"title":"ok"}}',
                xlsx_filename="entrada.xlsx",
            )

        root_mock.assert_called_once_with()
        cfg_mock.assert_called_once_with(Path("/repo"))
        build_mock.assert_called_once_with(
            repo_root=Path("/repo"),
            cfg={"cfg": True},
            xlsx_bytes=b"xlsx-bytes",
            xlsx_filename="entrada.xlsx",
            llm_payload={"response": {"title": "ok"}},
        )
        self.assertEqual(pptx_bytes, b"pptx-bytes")
        self.assertEqual(filename, "relatorio.pptx")
        self.assertIs(result, fake_result)

    def test_compose_presentation_from_inputs_serializes_build_result(self):
        fake_result = self._fake_build_result("api-output.pptx")

        with patch.object(
            build_pptx,
            "_build_presentation_artifact",
            return_value=(b"pptx-bytes", "api-output.pptx", fake_result),
        ) as artifact_mock:
            body = build_pptx.compose_presentation_from_inputs(
                b"xlsx-bytes",
                b'{"response":{}}',
                xlsx_filename="entrada.xlsx",
            )

        artifact_mock.assert_called_once_with(
            b"xlsx-bytes",
            b'{"response":{}}',
            xlsx_filename="entrada.xlsx",
        )
        self.assertEqual(body["filename"], "api-output.pptx")
        self.assertEqual(body["summary"]["generatedChartCount"], 4)
        self.assertTrue(body["pptxBase64"])

    def test_compose_presentation_download_response_returns_binary_file_response(self):
        fake_result = self._fake_build_result("download.pptx")

        with patch.object(
            build_pptx,
            "_build_presentation_artifact",
            return_value=(b"pptx-bytes", "download.pptx", fake_result),
        ) as artifact_mock:
            response = build_pptx.compose_presentation_download_response(
                b"xlsx-bytes",
                b'{"response":{}}',
                xlsx_filename="entrada.xlsx",
            )

        artifact_mock.assert_called_once_with(
            b"xlsx-bytes",
            b'{"response":{}}',
            xlsx_filename="entrada.xlsx",
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["Content-Type"], build_pptx.PPTX_CONTENT_TYPE)
        self.assertIn("download.pptx", response.headers["Content-Disposition"])
        self.assertEqual(response.get_data(), b"pptx-bytes")

    def test_read_upload_bytes_handles_supported_sources(self):
        self.assertEqual(build_pptx._read_upload_bytes(None), b"")
        self.assertEqual(build_pptx._read_upload_bytes(b"abc"), b"abc")
        self.assertEqual(build_pptx._read_upload_bytes(bytearray(b"abc")), b"abc")
        self.assertEqual(build_pptx._read_upload_bytes(_ContentUpload(b"content")), b"content")

        read_upload = _ReadUpload(b"stream")
        self.assertEqual(build_pptx._read_upload_bytes(read_upload), b"stream")
        self.assertEqual(read_upload.seek_calls, [0])

        seek_error_upload = _ReadUpload(b"stream", seek_raises=True)
        self.assertEqual(build_pptx._read_upload_bytes(seek_error_upload), b"stream")

    def test_read_upload_bytes_rejects_invalid_sources(self):
        with self.assertRaisesRegex(TypeError, "content invalido"):
            build_pptx._read_upload_bytes(_ContentUpload("not-bytes"))
        with self.assertRaisesRegex(TypeError, "conteudo nao binario"):
            build_pptx._read_upload_bytes(_ReadUpload("not-bytes"))
        with self.assertRaisesRegex(TypeError, "nao suportado"):
            build_pptx._read_upload_bytes(object())

    def test_upload_filename_uses_supported_attributes_and_basename(self):
        self.assertIsNone(build_pptx._upload_filename(None))
        self.assertEqual(
            build_pptx._upload_filename(_FilenameUpload(filename="/tmp/entrada.xlsx")),
            "entrada.xlsx",
        )
        self.assertEqual(
            build_pptx._upload_filename(_FilenameUpload(file_name="nested/entrada.xlsx")),
            "entrada.xlsx",
        )
        self.assertEqual(
            build_pptx._upload_filename(_FilenameUpload(original_filename="original.xlsx")),
            "original.xlsx",
        )
        self.assertEqual(
            build_pptx._upload_filename(_FilenameUpload(name="named.xlsx")),
            "named.xlsx",
        )
        self.assertIsNone(build_pptx._upload_filename(_FilenameUpload(filename="")))

    def test_first_value_skips_missing_empty_and_none_values(self):
        self.assertIsNone(build_pptx._first_value(None, "a"))
        self.assertIsNone(build_pptx._first_value({}, "a"))
        self.assertEqual(
            build_pptx._first_value({"a": None, "b": "", "c": 0}, "a", "b", "c"),
            0,
        )

    def test_decode_base64_field_handles_valid_empty_and_invalid_values(self):
        self.assertEqual(build_pptx._decode_base64_field(None, field_name="file"), b"")
        self.assertEqual(build_pptx._decode_base64_field("", field_name="file"), b"")
        self.assertEqual(
            build_pptx._decode_base64_field("eGxzeA==", field_name="file"),
            b"xlsx",
        )
        with self.assertRaisesRegex(ValueError, "deve ser string"):
            build_pptx._decode_base64_field(123, field_name="file")
        with self.assertRaisesRegex(ValueError, "base64 valido"):
            build_pptx._decode_base64_field("not base64", field_name="file")

    def test_json_to_bytes_handles_supported_and_invalid_values(self):
        self.assertEqual(build_pptx._json_to_bytes(None, field_name="llm"), b"")
        self.assertEqual(build_pptx._json_to_bytes("", field_name="llm"), b"")
        self.assertEqual(build_pptx._json_to_bytes(b"bytes", field_name="llm"), b"bytes")
        self.assertEqual(
            build_pptx._json_to_bytes(bytearray(b"bytes"), field_name="llm"),
            b"bytes",
        )
        self.assertEqual(build_pptx._json_to_bytes("texto", field_name="llm"), b"texto")
        self.assertEqual(
            build_pptx._json_to_bytes({"response": {"titulo": "olá"}}, field_name="llm"),
            '{"response": {"titulo": "olá"}}'.encode("utf-8"),
        )
        with self.assertRaisesRegex(ValueError, "serializado"):
            build_pptx._json_to_bytes({"bad": object()}, field_name="llm")

    def test_extract_request_payload_accepts_multipart_uploads(self):
        request_obj = _FakeRequest(
            files={
                "xlsx_file": _ContentUpload(b"xlsx-bytes", filename="/tmp/input.xlsx"),
                "llm_response_file": _ContentUpload(b'{"response":{}}'),
            }
        )

        self.assertEqual(
            build_pptx._extract_request_payload(request_obj),
            (b"xlsx-bytes", b'{"response":{}}', "input.xlsx"),
        )

    def test_extract_request_payload_accepts_multipart_with_form_llm(self):
        request_obj = _FakeRequest(
            files={"xlsx": _ContentUpload(b"xlsx-bytes", file_name="ignored.xlsx")},
            form={"llm_response_json": {"response": {"title": "ok"}}},
        )

        xlsx_bytes, llm_bytes, xlsx_filename = build_pptx._extract_request_payload(request_obj)

        self.assertEqual(xlsx_bytes, b"xlsx-bytes")
        self.assertEqual(llm_bytes, b'{"response": {"title": "ok"}}')
        self.assertEqual(xlsx_filename, "ignored.xlsx")

    def test_extract_request_payload_rejects_multipart_without_llm(self):
        request_obj = _FakeRequest(files={"xlsx_file": _ContentUpload(b"xlsx")})

        with self.assertRaisesRegex(ValueError, "Envie o JSON da LLM"):
            build_pptx._extract_request_payload(request_obj)

    def test_extract_request_payload_accepts_json_base64_payload(self):
        request_obj = _FakeRequest(
            json_payload={
                "xlsxBase64": "eGxzeC1ieXRlcw==",
                "llmResponseBase64": "eyJyZXNwb25zZSI6e319",
                "xlsxFilename": "nested/input.xlsx",
            }
        )

        self.assertEqual(
            build_pptx._extract_request_payload(request_obj),
            (b"xlsx-bytes", b'{"response":{}}', "input.xlsx"),
        )

    def test_extract_request_payload_accepts_json_inline_payload_and_typeerror_fallback(self):
        request_obj = _FakeRequest(
            json_payload={
                "xlsx_base64": "eGxzeA==",
                "llmResponse": {"response": {"ok": True}},
            },
            type_error_on_silent=True,
        )

        self.assertEqual(
            build_pptx._extract_request_payload(request_obj),
            (b"xlsx", b'{"response": {"ok": true}}', None),
        )

    def test_extract_request_payload_rejects_incomplete_json_and_unsupported_request(self):
        with self.assertRaisesRegex(ValueError, "No corpo JSON"):
            build_pptx._extract_request_payload(
                _FakeRequest(json_payload={"xlsxBase64": "eGxzeA=="})
            )
        with self.assertRaisesRegex(ValueError, "Requisicao sem payload suportado"):
            build_pptx._extract_request_payload(_FakeRequest())

    def test_handle_build_pptx_request_uses_global_request_when_omitted(self):
        fake_request = _FakeRequest(
            json_payload={
                "xlsxBase64": "eGxzeA==",
                "llmResponseBase64": "eyJyZXNwb25zZSI6e319",
                "filename": "input.xlsx",
            }
        )

        with (
            patch.object(build_pptx, "request", fake_request),
            patch.object(
                build_pptx,
                "compose_presentation_download_response",
                return_value="download-response",
            ) as compose_mock,
        ):
            response = build_pptx.handle_build_pptx_request()

        self.assertEqual(response, "download-response")
        compose_mock.assert_called_once_with(
            b"xlsx",
            b'{"response":{}}',
            xlsx_filename="input.xlsx",
        )

    def test_handle_build_pptx_request_returns_400_for_value_error(self):
        with (
            patch.object(
                build_pptx,
                "_extract_request_payload",
                side_effect=ValueError("bad"),
            ),
            patch.object(build_pptx.logging, "warning") as warning_mock,
        ):
            body, status = build_pptx.handle_build_pptx_request(_FakeRequest())

        self.assertEqual(status, 400)
        self.assertEqual(body["errorCode"], "invalid_request")
        warning_mock.assert_called_once()

    def test_handle_build_pptx_request_returns_500_for_unexpected_error(self):
        with (
            patch.object(
                build_pptx,
                "_extract_request_payload",
                return_value=(b"xlsx", b'{"response":{}}', None),
            ),
            patch.object(
                build_pptx,
                "compose_presentation_download_response",
                side_effect=RuntimeError("boom"),
            ),
            patch.object(build_pptx.logging, "exception") as exception_mock,
        ):
            body, status = build_pptx.handle_build_pptx_request(_FakeRequest())

        self.assertEqual(status, 500)
        self.assertEqual(body["errorCode"], "build_failed")
        exception_mock.assert_called_once()

    def test_build_pptx_resource_post_delegates_to_handler(self):
        with patch.object(
            build_pptx,
            "handle_build_pptx_request",
            return_value="ok-response",
        ) as handler_mock:
            response = build_pptx.BuildPptxResource().post()

        self.assertEqual(response, "ok-response")
        handler_mock.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
