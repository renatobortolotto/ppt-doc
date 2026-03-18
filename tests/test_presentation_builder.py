import types
import unittest
from pathlib import Path
from unittest.mock import patch

from presentation_builder import (
    ChartGenerationFailure,
    ChartGenerationResult,
    ChartGeneratorSpec,
    TextFieldExtractionResult,
    TextFieldFailure,
    build_presentation,
    build_text_mapping,
    build_text_mapping_with_failures,
    generate_chart_assets,
    load_llm_mapping_from_payload,
)


class TestPresentationBuilder(unittest.TestCase):
    def setUp(self):
        self.repo_root = Path(__file__).resolve().parents[1]
        self.cfg = {
            "pptx_template": "teste-design.gerado.updated.pptx",
            "pptx_output": "main_testing.pptx",
            "images_dir": ".",
            "allow_placeholder_text": False,
            "text_fields_config": "config/text_fields.json",
        }

    def test_load_llm_mapping_from_payload_accepts_wrapper(self):
        payload = {
            "response": {
                "titles": {"slide1_title": "Titulo"},
                "subtitles": {"slide1_subtitle": "Subtitulo"},
            }
        }
        mapping = load_llm_mapping_from_payload(payload)
        self.assertEqual(mapping["slide1_title"], "Titulo")
        self.assertEqual(mapping["slide1_subtitle"], "Subtitulo")

    def test_build_text_mapping_merges_filtered_llm_fields(self):
        llm_payload = {
            "response": {
                "titles": {
                    "slide1_title": "Titulo LLM",
                    "ignored_title": "Ignorado",
                },
                "subtitles": {"slide1_subtitle": "Subtitulo LLM"},
            }
        }
        with patch(
            "presentation_builder.extract_xlsx_to_text_mapping_tolerant",
            return_value=TextFieldExtractionResult(
                mapping={"ROE_RECORRENTE": "10,5%"},
                failures=(),
            ),
        ):
            mapping = build_text_mapping(
                repo_root=self.repo_root,
                cfg=self.cfg,
                xlsx_path=self.repo_root / "testing.xlsx",
                llm_payload=llm_payload,
            )

        self.assertEqual(mapping["ROE_RECORRENTE"], "10,5%")
        self.assertEqual(mapping["slide1_title"], "Titulo LLM")
        self.assertNotIn("ignored_title", mapping)

    def test_build_presentation_passes_text_payload_to_update(self):
        fake_result = (3, 0, 2, [], [], ["slide1_title", "slide1_subtitle"])
        with patch(
            "presentation_builder._load_validated_workbook",
        ) as load_workbook_mock:
            load_workbook_mock.return_value = types.SimpleNamespace(close=lambda: None)
            with patch(
                "presentation_builder.generate_chart_assets",
                return_value=ChartGenerationResult(generated_files=(Path("chart.png"),), failures=()),
            ):
                with patch(
                    "presentation_builder.build_text_mapping_with_failures",
                    return_value=TextFieldExtractionResult(
                        mapping={"slide1_title": "Titulo"},
                        failures=(),
                    ),
                ):
                    with patch(
                        "presentation_builder.update_presentation",
                        return_value=fake_result,
                    ) as update_mock:
                        result = build_presentation(
                            repo_root=self.repo_root,
                            cfg=self.cfg,
                            xlsx_path=self.repo_root / "testing.xlsx",
                            llm_payload={"response": {"titles": {"slide1_title": "Titulo"}}},
                        )

        self.assertEqual(result.generated_chart_count, 1)
        self.assertEqual(result.replaced_pictures, 3)
        self.assertEqual(result.replaced_text, 2)
        self.assertEqual(result.chart_failures, ())
        self.assertEqual(result.text_field_failures, ())
        update_mock.assert_called_once()

    def test_build_presentation_raises_clear_error_for_invalid_xlsx(self):
        with patch(
            "presentation_builder._load_validated_workbook",
            side_effect=ValueError("Arquivo enviado não é um XLSX válido"),
        ):
            with self.assertRaisesRegex(ValueError, "Arquivo Excel invalido"):
                build_presentation(
                    repo_root=self.repo_root,
                    cfg=self.cfg,
                    xlsx_path=self.repo_root / "testing.xlsx",
                )

    def test_generate_chart_assets_logs_failures_and_continues(self):
        def _ok_generator(*, xlsx_path: Path, output_dir: Path):
            return [output_dir / "ok.png"]

        def _bad_generator(*, xlsx_path: Path, output_dir: Path):
            raise ValueError("Range inesperado")

        specs = (
            ChartGeneratorSpec(
                key="ok",
                label="slide ok",
                generator=_ok_generator,
                output_files=("ok.png",),
            ),
            ChartGeneratorSpec(
                key="bad",
                label="slide bad",
                generator=_bad_generator,
                output_files=("bad1.png", "bad2.png"),
            ),
        )

        with patch("presentation_builder._chart_generators", return_value=specs):
            with patch("presentation_builder.logging.exception") as log_mock:
                result = generate_chart_assets(
                    xlsx_path=self.repo_root / "testing.xlsx",
                    images_dir=self.repo_root,
                )

        self.assertEqual(result.generated_files, (self.repo_root / "ok.png",))
        self.assertEqual(len(result.failures), 1)
        self.assertEqual(result.failures[0].generator_key, "bad")
        self.assertEqual(result.failures[0].output_files, ("bad1.png", "bad2.png"))
        log_mock.assert_called_once()

    def test_build_presentation_keeps_running_when_some_charts_fail(self):
        fake_result = (3, 0, 2, [], [], ["slide1_title"])
        chart_failures = (
            ChartGenerationFailure(
                generator_key="slide7",
                label="slide 7",
                output_files=("01_lucro_trimestres.png", "02_lucro_9m.png"),
                error="Valor nao numerico",
            ),
        )
        with patch(
            "presentation_builder._load_validated_workbook",
        ) as load_workbook_mock:
            load_workbook_mock.return_value = types.SimpleNamespace(close=lambda: None)
            with patch(
                "presentation_builder.generate_chart_assets",
                return_value=ChartGenerationResult(
                    generated_files=(Path("03_roe_trimestres.png"),),
                    failures=chart_failures,
                ),
            ):
                with patch(
                    "presentation_builder.build_text_mapping_with_failures",
                    return_value=TextFieldExtractionResult(
                        mapping={"slide1_title": "Titulo"},
                        failures=(),
                    ),
                ):
                    with patch(
                        "presentation_builder.update_presentation",
                        return_value=fake_result,
                    ):
                        result = build_presentation(
                            repo_root=self.repo_root,
                            cfg=self.cfg,
                            xlsx_path=self.repo_root / "testing.xlsx",
                        )

        self.assertEqual(result.generated_chart_count, 1)
        self.assertEqual(result.chart_failures, chart_failures)

    def test_build_text_mapping_with_failures_preserves_partial_texts(self):
        llm_payload = {
            "response": {
                "titles": {"slide1_title": "Titulo LLM"},
            }
        }
        failures = (
            TextFieldFailure(
                field_id="ROE_RECORRENTE",
                sheet="DRE Saida",
                a1_range="K20",
                error="Aba nao encontrada",
            ),
        )
        with patch(
            "presentation_builder.extract_xlsx_to_text_mapping_tolerant",
            return_value=TextFieldExtractionResult(
                mapping={"LL_RECORRENTE": "123"},
                failures=failures,
            ),
        ):
            result = build_text_mapping_with_failures(
                repo_root=self.repo_root,
                cfg=self.cfg,
                xlsx_path=self.repo_root / "testing.xlsx",
                llm_payload=llm_payload,
            )

        self.assertEqual(result.mapping["LL_RECORRENTE"], "123")
        self.assertEqual(result.mapping["slide1_title"], "Titulo LLM")
        self.assertEqual(result.failures, failures)

    def test_build_presentation_keeps_running_when_some_text_fields_fail(self):
        fake_result = (3, 0, 2, [], [], ["slide1_title"])
        text_failures = (
            TextFieldFailure(
                field_id="ROE_RECORRENTE",
                sheet="DRE Saida",
                a1_range="K20",
                error="Aba nao encontrada",
            ),
        )
        with patch(
            "presentation_builder._load_validated_workbook",
        ) as load_workbook_mock:
            load_workbook_mock.return_value = types.SimpleNamespace(close=lambda: None)
            with patch(
                "presentation_builder.generate_chart_assets",
                return_value=ChartGenerationResult(
                    generated_files=(Path("03_roe_trimestres.png"),),
                    failures=(),
                ),
            ):
                with patch(
                    "presentation_builder.build_text_mapping_with_failures",
                    return_value=TextFieldExtractionResult(
                        mapping={"slide1_title": "Titulo"},
                        failures=text_failures,
                    ),
                ):
                    with patch(
                        "presentation_builder.update_presentation",
                        return_value=fake_result,
                    ):
                        result = build_presentation(
                            repo_root=self.repo_root,
                            cfg=self.cfg,
                            xlsx_path=self.repo_root / "testing.xlsx",
                        )

        self.assertEqual(result.generated_chart_count, 1)
        self.assertEqual(result.text_field_failures, text_failures)


if __name__ == "__main__":
    unittest.main()
