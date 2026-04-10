import os
import sys
import types
import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

from presentation_builder import ChartGenerationFailure, TextFieldFailure


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
    def test_parse_llm_payload_rejects_empty_bytes(self):
        module = _load_main_framework()
        with self.assertRaisesRegex(ValueError, "JSON da LLM vazio"):
            module._parse_llm_payload(b"")

    def test_parse_llm_payload_rejects_invalid_utf8(self):
        module = _load_main_framework()
        with self.assertRaisesRegex(ValueError, "UTF-8"):
            module._parse_llm_payload(b"\xff")

    def test_parse_llm_payload_rejects_invalid_json(self):
        module = _load_main_framework()
        with self.assertRaisesRegex(ValueError, "JSON da LLM invalido"):
            module._parse_llm_payload(b"{")

    def test_compose_presentation_from_inputs_success(self):
        module = _load_main_framework()
        fake_build = types.SimpleNamespace(
            output_path=Path("/tmp/final.pptx"),
            replaced_pictures=3,
            replaced_placeholders=0,
            replaced_text=2,
            generated_chart_count=11,
            chart_failures=(
                ChartGenerationFailure(
                    generator_key="slide4",
                    label='slide 4 <script>alert("x")</script>',
                    output_files=("10_pizza_carteira.png",),
                    error='Valor nao numerico <img src=x onerror=alert(1)>',
                ),
            ),
            text_field_failures=(
                TextFieldFailure(
                    field_id="ROE_RECORRENTE",
                    sheet='DRE Saida <b>x</b>',
                    a1_range="K20",
                    error='Aba nao encontrada <script>alert(1)</script>',
                ),
            ),
            applied_text_keys=('slide1_title"><script>alert(1)</script>',),
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
        self.assertEqual(resp["summary"]["chartFailureCount"], 1)
        self.assertEqual(resp["summary"]["chartFailures"][0]["generatorKey"], "slide4")
        self.assertIn("&lt;script&gt;", resp["summary"]["chartFailures"][0]["label"])
        self.assertIn("&lt;img", resp["summary"]["chartFailures"][0]["error"])
        self.assertEqual(resp["summary"]["textFieldFailureCount"], 1)
        self.assertEqual(resp["summary"]["textFieldFailures"][0]["fieldId"], "ROE_RECORRENTE")
        self.assertIn("&lt;b&gt;", resp["summary"]["textFieldFailures"][0]["sheet"])
        self.assertIn("&lt;script&gt;", resp["summary"]["textFieldFailures"][0]["error"])
        self.assertIn("&lt;script&gt;", resp["summary"]["appliedTextKeys"][0])
        self.assertTrue(resp["pptxBase64"])

    def test_compose_presentation_from_inputs_prefers_api_output_filename(self):
        module = _load_main_framework()
        fake_build = types.SimpleNamespace(
            output_path=Path("/tmp/final.pptx"),
            replaced_pictures=0,
            replaced_placeholders=0,
            replaced_text=0,
            generated_chart_count=0,
            chart_failures=(),
            text_field_failures=(),
            applied_text_keys=(),
        )
        with patch(
            f"{module.__name__}.load_job_config",
            return_value={
                "api_output_filename": "nested/empresa.apresentacao.pptx",
                "pptx_output": "fallback.pptx",
            },
        ):
            with patch(
                f"{module.__name__}.build_presentation_from_bytes",
                return_value=(b"pptx-bytes", fake_build),
            ):
                resp = module.compose_presentation_from_inputs(
                    b"xlsx-bytes",
                    b'{"response":{"titles":{"slide1_title":"Titulo"}}}',
                )

        self.assertEqual(resp["filename"], "empresa.apresentacao.pptx")

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
        self.assertEqual(resp["errorCode"], "invalid_request")

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

    def test_compose_presentation_files_wraps_runtime_failures(self):
        module = _load_main_framework()
        with patch.object(
            module,
            "compose_presentation_from_inputs",
            side_effect=RuntimeError('<script>alert("boom")</script>'),
        ):
            resp = module.compose_presentation_files(
                DummyFileInput(b"xlsx"),
                DummyFileInput(b'{"response":{}}'),
            )

        self.assertEqual(
            resp["error"],
            "Falha ao montar o PowerPoint a partir do XLSX e do JSON da LLM.",
        )
        self.assertEqual(resp["errorCode"], "build_failed")
        self.assertEqual(resp["details"], "Consulte os logs do servidor para detalhes tecnicos.")


if __name__ == "__main__":
    unittest.main()
