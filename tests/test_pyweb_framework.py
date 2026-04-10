import types
import unittest
from unittest.mock import patch

from src.application.build_pptx import BuildPptxResource, handle_build_pptx_request
from src.controller.app import app
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


class TestPyWebFramework(unittest.TestCase):
    def test_create_routes_registers_build_pptx_resource(self):
        app.registered_routes.clear()

        create_routes()

        self.assertEqual(len(app.registered_routes), 1)
        self.assertEqual(app.registered_routes[0]["path"], "/api/build-pptx")
        self.assertIs(app.registered_routes[0]["handler"], BuildPptxResource)
        self.assertTrue(issubclass(app.registered_routes[0]["handler"], Resource))

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

    def test_handle_build_pptx_request_returns_400_for_missing_payload(self):
        body, status = handle_build_pptx_request(_FakeRequest())

        self.assertEqual(status, 400)
        self.assertIn("Requisicao invalida", body["error"])

    def test_main_module_imports_without_corporate_libs(self):
        with patch("src.controller.app.jwt_middleware", return_value=None):
            import main  # noqa: F401

        self.assertTrue(hasattr(main, "app"))


if __name__ == "__main__":
    unittest.main()
