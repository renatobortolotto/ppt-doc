import os
import sys
import types
import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch


class DummyFileInput:
    def __init__(self, content: bytes):
        self.content = content


def _install_dummy_modules():
    m_decorators = types.ModuleType("genai_framework.decorators")

    def file_input_route(_name):
        def decorator(fn):
            return fn
        return decorator

    m_decorators.file_input_route = file_input_route
    m_models = types.ModuleType("genai_framework.models")

    class FileInput:
        def __init__(self, content: bytes):
            self.content = content

    m_models.FileInput = FileInput

    sys.modules["genai_framework"] = types.ModuleType("genai_framework")
    sys.modules["genai_framework.decorators"] = m_decorators
    sys.modules["genai_framework.models"] = m_models


def _load_main_framework():
    _install_dummy_modules()
    path = Path(__file__).resolve().parents[1] / "main-framework.py"
    spec = importlib.util.spec_from_file_location("main_framework_module", str(path))
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    sys.modules["main_framework_module"] = module
    spec.loader.exec_module(module)
    return module


class TestMainFramework(unittest.TestCase):
    def test_compose_presentation_from_inputs_success(self):
        module = _load_main_framework()
        fake_build = types.SimpleNamespace(
            output_path=Path("/tmp/final.pptx"),
            replaced_pictures=3,
            replaced_placeholders=0,
            replaced_text=2,
            generated_chart_count=11,
            applied_text_keys=("slide1_title",),
        )
        with patch(
            f"{module.__name__}.load_job_config",
            return_value={"pptx_output": "main_testing.pptx"},
        ):
            with patch(
                f"{module.__name__}.build_presentation_from_bytes",
                return_value=(b"pptx-bytes", fake_build),
            ):
                resp = module.compose_presentation_from_inputs(
                    b"xlsx-bytes",
                    b'{"response":{"titles":{"slide1_title":"Titulo"}}}',
                )

        self.assertEqual(resp["filename"], "main_testing.pptx")
        self.assertEqual(resp["summary"]["replacedPictures"], 3)
        self.assertTrue(resp["pptxBase64"])

    def test_compose_presentation_files_invalid_json(self):
        module = _load_main_framework()
        resp = module.compose_presentation(
            DummyFileInput(b"xlsx"),
            DummyFileInput(b"not-json"),
        )
        self.assertIn("error", resp)
        self.assertEqual(
            resp["error"],
            "Falha ao montar o PowerPoint a partir do XLSX e do JSON da LLM.",
        )
        self.assertIn("details", resp)

    def test_compose_presentation_files_success(self):
        module = _load_main_framework()
        with patch.object(
            module,
            "compose_presentation_from_inputs",
            return_value={"filename": "main_testing.pptx", "pptxBase64": "b2s="},
        ) as compose_mock:
            resp = module.compose_presentation(
                DummyFileInput(b"xlsx"),
                DummyFileInput(b'{"response":{}}'),
            )

        self.assertEqual(resp["filename"], "main_testing.pptx")
        compose_mock.assert_called_once_with(b"xlsx", b'{"response":{}}')


if __name__ == "__main__":
    unittest.main()
