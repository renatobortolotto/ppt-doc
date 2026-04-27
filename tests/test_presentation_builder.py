import html
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from src.utils.xlsx_text_fields import TextFieldSpec

from presentation_builder import (
    BuildPresentationResult,
    ChartGenerationFailure,
    ChartGenerationResult,
    ChartGeneratorSpec,
    TextFieldExtractionResult,
    TextFieldFailure,
    build_presentation,
    build_presentation_from_bytes,
    build_text_mapping,
    extract_period_token,
    build_text_mapping_with_failures,
    generate_chart_assets,
    load_job_config,
    load_llm_mapping_from_payload,
    output_filename_for_xlsx,
    resolve_path,
    _chart_generators,
    _persist_generated_chart_files,
    _select_chart_generators,
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

    def test_output_filename_for_xlsx_uses_quarter_token(self):
        self.assertEqual(extract_period_token("Saída RGR_4T25_RI.xlsx"), "4T25")
        self.assertEqual(
            output_filename_for_xlsx(
                "Saída RGR_4t25_RI.xlsx",
                fallback_filename="main_testing.pptx",
            ),
            "PPT_4T25.pptx",
        )
        self.assertEqual(
            output_filename_for_xlsx(
                "arquivo_sem_periodo.xlsx",
                fallback_filename="nested/main_testing.pptx",
            ),
            "main_testing.pptx",
        )

    def test_output_filename_for_xlsx_neutralizes_html_sensitive_fallback(self):
        self.assertEqual(
            output_filename_for_xlsx(
                None,
                fallback_filename='evil<script>alert(1).pptx',
            ),
            "evil_script_alert(1).pptx",
        )
        self.assertEqual(
            output_filename_for_xlsx(
                'evil<script>_4T25.xlsx',
                fallback_filename='fallback<script>.pptx',
            ),
            "PPT_4T25.pptx",
        )

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
                "presentation_builder._load_text_fields_config",
                return_value=(self.repo_root / "config" / "text_fields.json", "DRE Saida", ()),
            ):
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
        self.assertEqual(update_mock.call_args.kwargs["pp_field_ids"], ())
        self.assertEqual(update_mock.call_args.kwargs["xlsx_path"], self.repo_root / "testing.xlsx")

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

    def test_generate_chart_assets_logs_start_and_success(self):
        def _slide4(*, xlsx_path: Path, output_dir: Path):
            return [output_dir / "slide4.png"]

        specs = (
            ChartGeneratorSpec(
                key="slide4",
                label="slide 4",
                generator=_slide4,
                output_files=("slide4.png",),
            ),
        )

        with patch("presentation_builder._chart_generators", return_value=specs):
            with patch("presentation_builder.logging.info") as info_mock:
                result = generate_chart_assets(
                    xlsx_path=self.repo_root / "testing.xlsx",
                    images_dir=self.repo_root,
                )

        self.assertEqual(result.generated_files, (self.repo_root / "slide4.png",))
        self.assertEqual(info_mock.call_count, 2)
        self.assertIn("Iniciando geracao de graficos", info_mock.call_args_list[0].args[0])
        self.assertIn("concluida com sucesso", info_mock.call_args_list[1].args[0])

    def test_chart_generators_exclude_slides_3_7_and_10_from_flow(self):
        generator_keys = {spec.key for spec in _chart_generators()}

        self.assertNotIn("slide3", generator_keys)
        self.assertNotIn("slide7", generator_keys)
        self.assertNotIn("slide10", generator_keys)
        self.assertIn("slide4", generator_keys)
        self.assertIn("slide8", generator_keys)

    def test_generate_chart_assets_filters_requested_slides(self):
        called: list[str] = []

        def _slide4(*, xlsx_path: Path, output_dir: Path):
            called.append("slide4")
            return [output_dir / "slide4.png"]

        def _slide8(*, xlsx_path: Path, output_dir: Path):
            called.append("slide8")
            return [output_dir / "slide8.png"]

        def _slide10(*, xlsx_path: Path, output_dir: Path):
            called.append("slide10")
            return [output_dir / "slide10.png"]

        specs = (
            ChartGeneratorSpec(
                key="slide4",
                label="slide 4",
                generator=_slide4,
                output_files=("slide4.png",),
            ),
            ChartGeneratorSpec(
                key="slide8",
                label="slide 8",
                generator=_slide8,
                output_files=("slide8.png",),
            ),
            ChartGeneratorSpec(
                key="slide10",
                label="slide 10",
                generator=_slide10,
                output_files=("slide10.png",),
            ),
        )

        with patch("presentation_builder._chart_generators", return_value=specs):
            result = generate_chart_assets(
                xlsx_path=self.repo_root / "testing.xlsx",
                images_dir=self.repo_root,
                only_slides=(4, 8, 10),
            )

        self.assertEqual(called, ["slide4", "slide8", "slide10"])
        self.assertEqual(
            result.generated_files,
            (
                self.repo_root / "slide4.png",
                self.repo_root / "slide8.png",
                self.repo_root / "slide10.png",
            ),
        )

    def test_build_presentation_keeps_running_when_some_charts_fail(self):
        fake_result = (3, 0, 2, [], [], ["slide1_title"])
        chart_failures = (
            ChartGenerationFailure(
                generator_key="slide4",
                label="slide 4",
                output_files=("10_pizza_carteira.png", "11_pizza_trimestres.png"),
                error="Valor nao numerico",
            ),
        )
        with patch(
            "presentation_builder._load_validated_workbook",
        ) as load_workbook_mock:
            load_workbook_mock.return_value = types.SimpleNamespace(close=lambda: None)
            with patch(
                "presentation_builder._load_text_fields_config",
                return_value=(self.repo_root / "config" / "text_fields.json", "DRE Saida", ()),
            ):
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

    def test_build_presentation_uses_isolated_images_dir_for_only_slides(self):
        fake_result = (3, 0, 2, [], [], ["slide1_title"])
        captured: dict[str, object] = {}

        def _fake_generate_chart_assets(*, xlsx_path: Path, images_dir: Path, only_slides=None):
            captured["generate_images_dir"] = images_dir
            captured["generate_only_slides"] = only_slides
            return ChartGenerationResult(generated_files=(images_dir / "slide8.png",), failures=())

        def _fake_update_presentation(*, pptx_path, output_path, images_dir, allow_placeholder_text, text_json, xlsx_path, text_payload, pp_field_ids):
            captured["update_images_dir"] = images_dir
            captured["update_pp_field_ids"] = pp_field_ids
            captured["update_xlsx_path"] = xlsx_path
            return fake_result

        def _fake_persist_generated_chart_files(*, generated_files, target_dir):
            captured["persist_generated_files"] = generated_files
            captured["persist_target_dir"] = target_dir
            return tuple(target_dir / Path(path).name for path in generated_files)

        with patch(
            "presentation_builder._load_validated_workbook",
        ) as load_workbook_mock:
            load_workbook_mock.return_value = types.SimpleNamespace(close=lambda: None)
            with patch(
                "presentation_builder._load_text_fields_config",
                return_value=(
                    self.repo_root / "config" / "text_fields.json",
                    "DRE Saida",
                    (TextFieldSpec(id="VAR_TEST", a1_range="A1", sheet="S", is_pp=True),),
                ),
            ):
                with patch(
                    "presentation_builder.generate_chart_assets",
                    side_effect=_fake_generate_chart_assets,
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
                            side_effect=_fake_update_presentation,
                        ):
                            with patch(
                                "presentation_builder._persist_generated_chart_files",
                                side_effect=_fake_persist_generated_chart_files,
                            ):
                                build_presentation(
                                    repo_root=self.repo_root,
                                    cfg=self.cfg,
                                    xlsx_path=self.repo_root / "testing.xlsx",
                                    only_slides=(1, 8),
                                )

        self.assertEqual(captured["generate_only_slides"], (1, 8))
        self.assertEqual(captured["generate_images_dir"], captured["update_images_dir"])
        self.assertNotEqual(captured["generate_images_dir"], self.repo_root)
        self.assertEqual(captured["update_pp_field_ids"], ("VAR_TEST",))
        self.assertEqual(captured["update_xlsx_path"], self.repo_root / "testing.xlsx")
        self.assertEqual(captured["persist_generated_files"], (captured["generate_images_dir"] / "slide8.png",))
        self.assertEqual(captured["persist_target_dir"], self.repo_root)

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
                "presentation_builder._load_text_fields_config",
                return_value=(self.repo_root / "config" / "text_fields.json", "DRE Saida", ()),
            ):
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

    def test_build_presentation_escapes_result_metadata_before_return(self):
        payloads = (
            "<script>alert(1)</script>",
            '"><img src=x onerror=alert(1)>',
            "' onmouseover='alert(1)",
            "</script><script>alert(1)</script>",
            "<svg/onload=alert(1)>",
        )
        fake_result = (
            3,
            0,
            2,
            [],
            [],
            payloads,
        )
        chart_failures = (
            ChartGenerationFailure(
                generator_key=payloads[1],
                label=payloads[0],
                output_files=(payloads[2],),
                error=payloads[3],
            ),
        )
        text_failures = (
            TextFieldFailure(
                field_id=payloads[1],
                sheet=payloads[2],
                a1_range=payloads[3],
                error=payloads[4],
            ),
        )

        with (
            patch("presentation_builder._load_validated_workbook") as load_workbook_mock,
            patch(
                "presentation_builder._load_text_fields_config",
                return_value=(self.repo_root / "config" / "text_fields.json", "DRE Saida", ()),
            ),
            patch(
                "presentation_builder.generate_chart_assets",
                return_value=ChartGenerationResult(
                    generated_files=(Path("03_roe_trimestres.png"),),
                    failures=chart_failures,
                ),
            ),
            patch(
                "presentation_builder.build_text_mapping_with_failures",
                return_value=TextFieldExtractionResult(
                    mapping={"slide1_title": "Titulo"},
                    failures=text_failures,
                ),
            ),
            patch(
                "presentation_builder.update_presentation",
                return_value=fake_result,
            ),
        ):
            load_workbook_mock.return_value = types.SimpleNamespace(close=lambda: None)
            result = build_presentation(
                repo_root=self.repo_root,
                cfg=self.cfg,
                xlsx_path=self.repo_root / "testing.xlsx",
            )

        rendered = " ".join(
            (
                result.chart_failures[0].generator_key,
                result.chart_failures[0].label,
                result.chart_failures[0].output_files[0],
                result.chart_failures[0].error,
                result.text_field_failures[0].field_id,
                result.text_field_failures[0].sheet or "",
                result.text_field_failures[0].a1_range,
                result.text_field_failures[0].error,
                " ".join(result.applied_text_keys),
            )
        )
        for payload in payloads:
            self.assertNotIn(payload, rendered)
            self.assertIn(html.escape(str(payload), quote=True), rendered)
        self.assertNotIn("<script>", rendered)
        self.assertNotIn("<img", rendered)
        self.assertNotIn("<svg", rendered)

    def test_resolve_path_supports_relative_and_absolute_inputs(self):
        relative = resolve_path(self.repo_root, "config/job_config.json")
        absolute_input = Path("/tmp/ppt-doc-absolute.json")
        absolute = resolve_path(self.repo_root, str(absolute_input))

        self.assertEqual(relative, (self.repo_root / "config" / "job_config.json").resolve())
        self.assertEqual(absolute, absolute_input)

    def test_resolve_path_escapes_html_sensitive_input(self):
        payloads = (
            "<script>alert(1)</script>",
            "\"><img src=x onerror=alert(1)>",
            "' onmouseover='alert(1)",
            "</script><script>alert(1)</script>",
            "<svg/onload=alert(1)>",
        )

        for payload in payloads:
            raw_path = f"reports/{payload}.pptx"
            resolved = resolve_path(self.repo_root, raw_path)
            rendered_path = str(resolved)

            self.assertNotIn(raw_path, rendered_path)
            self.assertIn(html.escape(raw_path, quote=True), rendered_path)
            self.assertNotIn("<", rendered_path)
            self.assertNotIn(">", rendered_path)
            self.assertNotIn('"', rendered_path)
            self.assertNotIn("'", rendered_path)

    def test_load_job_config_rejects_non_object_json(self):
        with tempfile.TemporaryDirectory() as td:
            repo_root = Path(td)
            config_dir = repo_root / "config"
            config_dir.mkdir(parents=True, exist_ok=True)
            (config_dir / "job_config.json").write_text("[]", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "objeto"):
                load_job_config(repo_root)

    def test_select_chart_generators_rejects_unavailable_slides(self):
        with self.assertRaisesRegex(ValueError, "Nenhum dos slides informados"):
            _select_chart_generators((999,))

    def test_select_chart_generators_rejects_removed_slides_3_7_and_10(self):
        with self.assertRaisesRegex(ValueError, "Nenhum dos slides informados"):
            _select_chart_generators((3, 7, 10))

    def test_persist_generated_chart_files_copies_into_target_dir(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            source_dir = tmpdir / "source"
            target_dir = tmpdir / "target"
            source_dir.mkdir(parents=True, exist_ok=True)
            target_dir.mkdir(parents=True, exist_ok=True)

            generated_file = source_dir / "chart.png"
            generated_file.write_bytes(b"png-bytes")

            persisted = _persist_generated_chart_files(
                generated_files=(generated_file,),
                target_dir=target_dir,
            )

            self.assertEqual(persisted, (target_dir / "chart.png",))
            self.assertEqual((target_dir / "chart.png").read_bytes(), b"png-bytes")

    def test_build_presentation_rejects_only_slides_with_skip_charts(self):
        with patch(
            "presentation_builder._load_validated_workbook",
        ) as load_workbook_mock:
            load_workbook_mock.return_value = types.SimpleNamespace(close=lambda: None)

            with self.assertRaisesRegex(ValueError, "only_slides nao pode ser usado junto com skip_charts"):
                build_presentation(
                    repo_root=self.repo_root,
                    cfg=self.cfg,
                    xlsx_path=self.repo_root / "testing.xlsx",
                    skip_charts=True,
                    only_slides=(8,),
                )

    def test_build_presentation_from_bytes_rejects_empty_xlsx(self):
        with self.assertRaisesRegex(ValueError, "XLSX vazio"):
            build_presentation_from_bytes(
                repo_root=self.repo_root,
                cfg=self.cfg,
                xlsx_bytes=b"",
            )

    def test_build_presentation_from_bytes_uses_api_output_filename(self):
        cfg = dict(self.cfg)
        cfg["api_output_filename"] = "nested/empresa-final.pptx"

        def _fake_build_presentation(
            *,
            repo_root,
            cfg,
            xlsx_path,
            llm_payload,
            output_path,
            images_dir,
            skip_charts,
        ):
            self.assertEqual(repo_root, self.repo_root)
            self.assertEqual(cfg["api_output_filename"], "nested/empresa-final.pptx")
            self.assertEqual(xlsx_path.read_bytes(), b"xlsx-bytes")
            self.assertEqual(images_dir.name, "images")
            self.assertEqual(llm_payload, {"response": {"titles": {"slide1_title": "Titulo"}}})
            self.assertFalse(skip_charts)
            output_path.write_bytes(b"pptx-bytes")
            return BuildPresentationResult(
                output_path=output_path,
                replaced_pictures=2,
                replaced_placeholders=1,
                replaced_text=3,
                generated_chart_count=4,
                chart_failures=(),
                text_field_failures=(),
                applied_text_keys=("slide1_title",),
            )

        with patch(
            "presentation_builder.build_presentation",
            side_effect=_fake_build_presentation,
        ):
            pptx_bytes, result = build_presentation_from_bytes(
                repo_root=self.repo_root,
                cfg=cfg,
                xlsx_bytes=b"xlsx-bytes",
                llm_payload={"response": {"titles": {"slide1_title": "Titulo"}}},
            )

        self.assertEqual(pptx_bytes, b"pptx-bytes")
        self.assertEqual(result.output_path, Path("empresa-final.pptx"))
        self.assertEqual(result.replaced_pictures, 2)
        self.assertEqual(result.replaced_placeholders, 1)
        self.assertEqual(result.replaced_text, 3)
        self.assertEqual(result.generated_chart_count, 4)
        self.assertEqual(result.applied_text_keys, ("slide1_title",))

    def test_build_presentation_from_bytes_uses_xlsx_period_in_output_filename(self):
        cfg = dict(self.cfg)

        def _fake_build_presentation(
            *,
            repo_root,
            cfg,
            xlsx_path,
            llm_payload,
            output_path,
            images_dir,
            skip_charts,
        ):
            self.assertEqual(output_path.name, "PPT_4T25.pptx")
            output_path.write_bytes(b"pptx-bytes")
            return BuildPresentationResult(
                output_path=output_path,
                replaced_pictures=0,
                replaced_placeholders=0,
                replaced_text=0,
                generated_chart_count=0,
                chart_failures=(),
                text_field_failures=(),
                applied_text_keys=(),
            )

        with patch(
            "presentation_builder.build_presentation",
            side_effect=_fake_build_presentation,
        ):
            pptx_bytes, result = build_presentation_from_bytes(
                repo_root=self.repo_root,
                cfg=cfg,
                xlsx_bytes=b"xlsx-bytes",
                xlsx_filename="Saída RGR_4T25_RI.xlsx",
            )

        self.assertEqual(pptx_bytes, b"pptx-bytes")
        self.assertEqual(result.output_path, Path("PPT_4T25.pptx"))


if __name__ == "__main__":
    unittest.main()
