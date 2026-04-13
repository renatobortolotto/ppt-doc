import types
import unittest
from pathlib import Path
from unittest.mock import patch

from src.application.build_pptx import (
    BuildPptxResource,
    compose_presentation_from_inputs,
    handle_build_pptx_request,
)
from src.controller.app import app
from src.infrastructure.config import prefix
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
    def __init__(self, *, files=None, form=None, json_payload=None):
        self.files = files or {}
        self.form = form or {}
        self._json_payload = json_payload
        self.headers = {}
        self.path = "/api/build-pptx"

    def get_json(self, silent: bool = False):
        return self._json_payload


def _expected_route_path() -> str:
    normalized_prefix = prefix.rstrip("/")
    return f"{normalized_prefix}/build-pptx" if normalized_prefix else "/build-pptx"


def _registered_build_pptx_handlers():
    full_path = _expected_route_path()

    registered_routes = getattr(app, "registered_routes", None)
    if isinstance(registered_routes, list):
        return [
            route["handler"]
            for route in registered_routes
            if isinstance(route, dict) and route.get("path") == full_path
        ]

    url_map = getattr(app, "url_map", None)
    view_functions = getattr(app, "view_functions", None)
    if url_map is None or view_functions is None:
        raise AssertionError("Aplicacao nao expoe registered_routes nem url_map/view_functions")

    handlers = []
    for rule in url_map.iter_rules():
        if rule.rule != full_path:
            continue
        view = view_functions.get(rule.endpoint)
        handlers.append(getattr(view, "view_class", view))
    return handlers


class TestPyWebFramework(unittest.TestCase):
    def test_create_routes_registers_build_pptx_resource(self):
        create_routes()
        handlers = _registered_build_pptx_handlers()

        self.assertGreaterEqual(len(handlers), 1)
        self.assertIs(handlers[0], BuildPptxResource)
        self.assertTrue(issubclass(handlers[0], Resource))

    def test_create_routes_is_idempotent(self):
        create_routes()
        create_routes()

        handlers = _registered_build_pptx_handlers()
        self.assertEqual(sum(handler is BuildPptxResource for handler in handlers), 1)

    def test_build_pptx_resource_post_delegates_to_request_handler(self):
        resource = BuildPptxResource()

        with patch(
            "src.application.build_pptx.handle_build_pptx_request",
            return_value=({"filename": "ok.pptx"}, 200),
        ) as handler_mock:
            body, status = resource.post()

        self.assertEqual(status, 200)
        self.assertEqual(body["filename"], "ok.pptx")
        handler_mock.assert_called_once_with()

    def test_handle_build_pptx_request_accepts_multipart_uploads(self):
        request_obj = _FakeRequest(
            files={
                "xlsx_file": _FakeUpload(b"xlsx-bytes"),
                "llm_response_file": _FakeUpload(b'{"response":{}}'),
            }
        )

        with patch(
            "src.application.build_pptx.compose_presentation_from_inputs",
            return_value={"filename": "ok.pptx"},
        ) as compose_mock:
            body, status = handle_build_pptx_request(request_obj)

        self.assertEqual(status, 200)
        self.assertEqual(body["filename"], "ok.pptx")
        compose_mock.assert_called_once_with(b"xlsx-bytes", b'{"response":{}}')

    def test_handle_build_pptx_request_accepts_json_base64_payload(self):
        request_obj = _FakeRequest(
            json_payload={
                "xlsxBase64": "eGxzeC1ieXRlcw==",
                "llmResponseBase64": "eyJyZXNwb25zZSI6e319",
            }
        )

        with patch(
            "src.application.build_pptx.compose_presentation_from_inputs",
            return_value={"filename": "ok.pptx"},
        ) as compose_mock:
            body, status = handle_build_pptx_request(request_obj)

        self.assertEqual(status, 200)
        self.assertEqual(body["filename"], "ok.pptx")
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
            "src.application.build_pptx.compose_presentation_from_inputs",
            side_effect=RuntimeError('<script>alert("x")</script>'),
        ):
            body, status = handle_build_pptx_request(request_obj)

        self.assertEqual(status, 500)
        self.assertEqual(body["errorCode"], "build_failed")
        self.assertNotIn("<script>", body["details"])

    def test_main_module_imports_without_corporate_libs(self):
        with patch("src.controller.app.jwt_middleware", return_value=None):
            import main  # noqa: F401

        self.assertTrue(hasattr(main, "app"))


if __name__ == "__main__":
    unittest.main()
