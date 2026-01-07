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


class DummyModels:
    def __init__(self):
        self.last_call = None

    def generate_content(self, model, contents, config):
        # Record for optional assertions
        self.last_call = {
            "model": model,
            "contents": contents,
            "config": config,
        }
        resp = types.SimpleNamespace()
        resp.text = '{"ok": true}'
        return resp


class DummyClient:
    def __init__(self, **kwargs):
        self.kwargs = kwargs
        self.models = DummyModels()


class DummyTypes:
    class GenerateContentConfig:
        def __init__(self, temperature: float, max_output_tokens: int):
            self.temperature = temperature
            self.max_output_tokens = max_output_tokens


def _install_dummy_modules():
    # genai_framework.decorators
    m_decorators = types.ModuleType("genai_framework.decorators")

    def file_input_route(_name):
        def decorator(fn):
            return fn
        return decorator

    m_decorators.file_input_route = file_input_route

    # genai_framework.models
    m_models = types.ModuleType("genai_framework.models")

    class FileInput:
        def __init__(self, content: bytes):
            self.content = content

    m_models.FileInput = FileInput

    # google.genai
    m_google_genai = types.ModuleType("google.genai")
    m_google_genai.Client = DummyClient
    m_google_genai.types = DummyTypes

    sys.modules["genai_framework"] = types.ModuleType("genai_framework")
    sys.modules["genai_framework.decorators"] = m_decorators
    sys.modules["genai_framework.models"] = m_models
    sys.modules["google.genai"] = m_google_genai


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
    def test_analyze_file_success(self):
        module = _load_main_framework()
        with patch(f"{module.__name__}.parse_specs_json", return_value=[{"id": "lucroTrimestre"}]):
            with patch(f"{module.__name__}.extract_xlsx_bytes_to_dict", return_value={"lucroTrimestre": {"labels": ["A"], "values": [1]}}):
                # Ensure env var is set to avoid path resolution noise
                os.environ["SPECS_JSON_PATH"] = str(Path("config/specs.json").resolve())
                resp = module.analyze_file(DummyFileInput(b"xlsx-bytes"))
        self.assertIn("response", resp)
        self.assertEqual(resp["response"], {"ok": True})

    def test_analyze_file_invalid_xlsx(self):
        module = _load_main_framework()
        with patch(f"{module.__name__}.parse_specs_json", return_value=[{"id": "x"}]):
            with patch(f"{module.__name__}.extract_xlsx_bytes_to_dict", side_effect=ValueError("Arquivo enviado não é um XLSX válido")):
                os.environ["SPECS_JSON_PATH"] = str(Path("config/specs.json").resolve())
                resp = module.analyze_file(DummyFileInput(b"bad"))
        self.assertIn("error", resp)
        self.assertEqual(resp["error"], "Falha ao extrair dados do XLSX usando specs.json.")
        self.assertIn("details", resp)


if __name__ == "__main__":
    unittest.main()
