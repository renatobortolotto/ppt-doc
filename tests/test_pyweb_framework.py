import importlib
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from src.application.build_pptx import (
    BuildPptxResource,
    PPTX_CONTENT_TYPE,
    compose_presentation_from_inputs,
    compose_presentation_download_response,
    handle_build_pptx_request,
)
from src.infrastructure.framework_compat import Resource
from src.routes import create_routes


class _FakeUpload:
    def __init__(self, data: bytes):
        self._data = data

    def read(self) -> bytes:
        return self._data

    def seek(self, _offset: int) -> None:
        return None


class _FakeRequest:
    def __init__(self, *, files=None, form=None, json_payload=None, method="POST"):
        self.files = files or {}
        self.form = form or {}
        self._json_payload = json_payload
        self.headers = {}
        self.path = "/api/build-pptx"
        self.method = method

    def get_json(self, silent: bool = False):
        return self._json_payload


class TestPyWebFramework(unittest.TestCase):
    def test_create_routes_registers_build_pptx_resource(self):
        with patch("src.routes._build_pptx_route_registered", return_value=False):
            with patch("src.routes.op_app.create_route") as create_route_mock:
                create_routes()

        create_route_mock.assert_called_once_with(BuildPptxResource, "/build-pptx")
        self.assertTrue(issubclass(BuildPptxResource, Resource))

    def test_create_routes_is_idempotent(self):
        with patch("src.routes._build_pptx_route_registered", side_effect=[False, True]):
            with patch("src.routes.op_app.create_route") as create_route_mock:
                create_routes()
                create_routes()

        create_route_mock.assert_called_once_with(BuildPptxResource, "/build-pptx")

    def test_build_pptx_resource_post_delegates_to_request_handler(self):
        resource = BuildPptxResource()

        with patch(
            "src.application.build_pptx.handle_build_pptx_request",
            return_value="ok-response",
        ) as handler_mock:
            response = resource.post()

        self.assertEqual(response, "ok-response")
        handler_mock.assert_called_once_with()

    def test_handle_build_pptx_request_accepts_multipart_uploads(self):
        request_obj = _FakeRequest(
            files={
                "xlsx_file": _FakeUpload(b"xlsx-bytes"),
                "llm_response_file": _FakeUpload(b'{"response":{}}'),
            }
        )

        with patch(
            "src.application.build_pptx.compose_presentation_download_response",
            return_value="download-response",
        ) as compose_mock:
            response = handle_build_pptx_request(request_obj)

        self.assertEqual(response, "download-response")
        compose_mock.assert_called_once_with(b"xlsx-bytes", b'{"response":{}}')

    def test_handle_build_pptx_request_accepts_json_base64_payload(self):
        request_obj = _FakeRequest(
            json_payload={
                "xlsxBase64": "eGxzeC1ieXRlcw==",
                "llmResponseBase64": "eyJyZXNwb25zZSI6e319",
            }
        )

        with patch(
            "src.application.build_pptx.compose_presentation_download_response",
            return_value="download-response",
        ) as compose_mock:
            response = handle_build_pptx_request(request_obj)

        self.assertEqual(response, "download-response")
        compose_mock.assert_called_once_with(b"xlsx-bytes", b'{"response":{}}')

    def test_compose_presentation_from_inputs_serializes_success_response(self):
        fake_build = types.SimpleNamespace(
            output_path=Path("saida.pptx"),
            replaced_pictures=1,
            replaced_placeholders=0,
            replaced_text=2,
            generated_chart_count=3,
            chart_failures=(),
            text_field_failures=(),
            applied_text_keys=("slide1_title",),
        )

        with patch(
            "src.application.build_pptx.load_job_config",
            return_value={"pptx_output": "main_testing.pptx"},
        ):
            with patch(
                "src.application.build_pptx.build_presentation_from_bytes",
                return_value=(b"pptx-bytes", fake_build),
            ):
                body = compose_presentation_from_inputs(
                    b"xlsx-bytes",
                    b'{"response":{}}',
                )

        self.assertEqual(body["filename"], "main_testing.pptx")
        self.assertEqual(body["summary"]["generatedChartCount"], 3)
        self.assertTrue(body["pptxBase64"])

    def test_compose_presentation_download_response_returns_binary_attachment(self):
        fake_build = types.SimpleNamespace(
            output_path=Path("saida.pptx"),
            replaced_pictures=1,
            replaced_placeholders=0,
            replaced_text=2,
            generated_chart_count=3,
            chart_failures=(),
            text_field_failures=(),
            applied_text_keys=("slide1_title",),
        )

        with patch(
            "src.application.build_pptx.load_job_config",
            return_value={"pptx_output": "main_testing.pptx"},
        ):
            with patch(
                "src.application.build_pptx.build_presentation_from_bytes",
                return_value=(b"pptx-bytes", fake_build),
            ):
                response = compose_presentation_download_response(
                    b"xlsx-bytes",
                    b'{"response":{}}',
                )

        self.assertEqual(getattr(response, "status_code", None), 200)
        self.assertEqual(response.headers.get("Content-Type"), PPTX_CONTENT_TYPE)
        self.assertIn("attachment;", response.headers.get("Content-Disposition", ""))
        self.assertIn("main_testing.pptx", response.headers.get("Content-Disposition", ""))
        self.assertEqual(response.get_data(), b"pptx-bytes")

    def test_handle_build_pptx_request_returns_400_for_missing_payload(self):
        body, status = handle_build_pptx_request(_FakeRequest())

        self.assertEqual(status, 400)
        self.assertIn("Requisicao invalida", body["error"])
        self.assertEqual(body["errorCode"], "invalid_request")

    def test_handle_build_pptx_request_hides_raw_processing_errors(self):
        request_obj = _FakeRequest(
            files={
                "xlsx_file": _FakeUpload(b"xlsx-bytes"),
                "llm_response_file": _FakeUpload(b'{"response":{}}'),
            }
        )

        with patch(
            "src.application.build_pptx.compose_presentation_download_response",
            side_effect=RuntimeError('<script>alert("x")</script>'),
        ):
            body, status = handle_build_pptx_request(request_obj)

        self.assertEqual(status, 500)
        self.assertEqual(body["errorCode"], "build_failed")
        self.assertNotIn("<script>", body["details"])

    def test_main_module_imports_without_corporate_libs(self):
        with patch("src.controller.app.jwt_middleware", return_value=None):
            main_module = importlib.import_module("main")

        self.assertIn("app", vars(main_module))


if __name__ == "__main__":
    unittest.main()
