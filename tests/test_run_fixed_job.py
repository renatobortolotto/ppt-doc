import json
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from presentation_builder import ChartGenerationFailure
from src.utils.xlsx_text_fields import TextFieldFailure
from run_fixed_job import _parse_only_slides
from run_fixed_job import main as run_fixed_job_main


class TestRunFixedJob(unittest.TestCase):
    def test_parse_only_slides_accepts_ranges_and_list(self):
        self.assertEqual(_parse_only_slides("1-3, 8, 10"), (1, 2, 3, 8, 10))

    def test_parse_only_slides_rejects_invalid_values(self):
        with self.assertRaises(ValueError):
            _parse_only_slides("3-a")

    def test_main_calls_builder_with_llm_json_and_only_slides(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            xlsx_path = tmpdir / "input.xlsx"
            xlsx_path.write_bytes(b"xlsx-placeholder")
            llm_json_path = tmpdir / "llm.json"
            llm_payload = {"response": {"titles": {"slide1_title": "Titulo"}}}
            llm_json_path.write_text(json.dumps(llm_payload), encoding="utf-8")

            fake_result = types.SimpleNamespace(
                output_path=tmpdir / "out.pptx",
                replaced_pictures=2,
                replaced_text=1,
                generated_chart_count=4,
                chart_failures=(),
                text_field_failures=(),
            )

            with patch.object(
                sys,
                "argv",
                [
                    "run_fixed_job.py",
                    "--xlsx",
                    str(xlsx_path),
                    "--llm-json",
                    str(llm_json_path),
                    "--only-slides",
                    "8-9",
                ],
            ):
                with patch("run_fixed_job._configure_logging") as configure_mock:
                    with patch(
                        "run_fixed_job._load_job_config",
                        return_value={"pptx_output": "main_testing.pptx"},
                    ):
                        with patch(
                            "run_fixed_job.build_presentation",
                            return_value=fake_result,
                        ) as build_mock:
                            run_fixed_job_main()

        configure_mock.assert_called_once()
        build_kwargs = build_mock.call_args.kwargs
        self.assertEqual(build_kwargs["cfg"], {"pptx_output": "main_testing.pptx"})
        self.assertEqual(build_kwargs["xlsx_path"], xlsx_path.resolve())
        self.assertEqual(build_kwargs["llm_payload"], llm_payload)
        self.assertEqual(build_kwargs["only_slides"], (8, 9))
        self.assertFalse(build_kwargs["skip_charts"])

    def test_main_logs_chart_and_text_failures(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            xlsx_path = tmpdir / "input.xlsx"
            xlsx_path.write_bytes(b"xlsx-placeholder")

            fake_result = types.SimpleNamespace(
                output_path=tmpdir / "out.pptx",
                replaced_pictures=2,
                replaced_text=1,
                generated_chart_count=4,
                chart_failures=(
                    ChartGenerationFailure(
                        generator_key="slide4",
                        label="slide 4",
                        output_files=("10_pizza_carteira.png",),
                        error="valor invalido",
                    ),
                ),
                text_field_failures=(
                    TextFieldFailure(
                        field_id="ROE_RECORRENTE",
                        sheet="DRE Saida",
                        a1_range="K20",
                        error="aba ausente",
                    ),
                ),
            )

            with patch.object(sys, "argv", ["run_fixed_job.py", "--xlsx", str(xlsx_path)]):
                with patch("run_fixed_job._configure_logging"):
                    with patch(
                        "run_fixed_job._load_job_config",
                        return_value={"pptx_output": "main_testing.pptx"},
                    ):
                        with patch(
                            "run_fixed_job.build_presentation",
                            return_value=fake_result,
                        ):
                            with patch("run_fixed_job.logging.warning") as warning_mock:
                                run_fixed_job_main()

        self.assertEqual(warning_mock.call_count, 2)
        self.assertIn("Grafico com falha", warning_mock.call_args_list[0].args[0])
        self.assertIn("Campo de texto com falha", warning_mock.call_args_list[1].args[0])

    def test_main_rejects_invalid_only_slides_argument(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            xlsx_path = tmpdir / "input.xlsx"
            xlsx_path.write_bytes(b"xlsx-placeholder")

            with patch.object(
                sys,
                "argv",
                ["run_fixed_job.py", "--xlsx", str(xlsx_path), "--only-slides", "3-a"],
            ):
                with patch("run_fixed_job._configure_logging"):
                    with patch(
                        "run_fixed_job._load_job_config",
                        return_value={"pptx_output": "main_testing.pptx"},
                    ):
                        with self.assertRaises(SystemExit):
                            run_fixed_job_main()


if __name__ == "__main__":
    unittest.main()
