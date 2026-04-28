import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from presentation_builder import ChartGenerationFailure, ChartGenerationResult
from scripts import generate_chart_assets as generate_chart_assets_script


class TestGenerateChartAssetsScript(unittest.TestCase):
    def test_main_uses_project_root_output_dir_and_prints_summary(self):
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            xlsx_path = td_path / "input.xlsx"
            xlsx_path.write_bytes(b"fake-xlsx")

            generated_file = td_path / "out" / "chart.png"
            fake_result = ChartGenerationResult(
                generated_files=(generated_file,),
                failures=(),
            )

            stdout = io.StringIO()
            with (
                patch.object(
                    generate_chart_assets_script,
                    "_load_job_config",
                    return_value={"images_dir": "../ignored"},
                ),
                patch.object(generate_chart_assets_script, "_configure_logging"),
                patch.object(
                    generate_chart_assets_script,
                    "generate_chart_assets",
                    return_value=fake_result,
                ) as generate_mock,
                patch("sys.argv", ["generate_chart_assets.py", "--xlsx", str(xlsx_path)]),
                redirect_stdout(stdout),
            ):
                exit_code = generate_chart_assets_script.main()

            self.assertEqual(exit_code, 0)
            generate_mock.assert_called_once()
            self.assertEqual(generate_mock.call_args.kwargs["xlsx_path"], xlsx_path.resolve())
            self.assertEqual(
                generate_mock.call_args.kwargs["images_dir"],
                generate_chart_assets_script.REPO_ROOT.resolve(),
            )
            self.assertIsNone(generate_mock.call_args.kwargs["only_slides"])

            summary = json.loads(stdout.getvalue())
            self.assertEqual(summary["generatedCount"], 1)
            self.assertEqual(summary["failureCount"], 0)
            self.assertEqual(summary["generatedFiles"], [str(generated_file)])

    def test_main_returns_one_in_strict_mode_when_any_generator_fails(self):
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            xlsx_path = td_path / "input.xlsx"
            xlsx_path.write_bytes(b"fake-xlsx")

            fake_result = ChartGenerationResult(
                generated_files=(),
                failures=(
                    ChartGenerationFailure(
                        generator_key="slide11",
                        label="slide 11",
                        output_files=("11_chart.png",),
                        error="aba nao encontrada",
                    ),
                ),
            )

            stdout = io.StringIO()
            with (
                patch.object(
                    generate_chart_assets_script,
                    "_load_job_config",
                    return_value={"images_dir": "."},
                ),
                patch.object(generate_chart_assets_script, "_configure_logging"),
                patch.object(
                    generate_chart_assets_script,
                    "generate_chart_assets",
                    return_value=fake_result,
                ),
                patch(
                    "sys.argv",
                    [
                        "generate_chart_assets.py",
                        "--xlsx",
                        str(xlsx_path),
                        "--only-slides",
                        "11,12",
                        "--strict",
                    ],
                ),
                redirect_stdout(stdout),
            ):
                exit_code = generate_chart_assets_script.main()

            self.assertEqual(exit_code, 1)
            summary = json.loads(stdout.getvalue())
            self.assertEqual(summary["failureCount"], 1)
            self.assertEqual(summary["failures"][0]["generatorKey"], "slide11")


if __name__ == "__main__":
    unittest.main()
