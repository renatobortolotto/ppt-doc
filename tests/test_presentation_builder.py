import types
import unittest
from pathlib import Path
from unittest.mock import patch

from presentation_builder import (
    build_presentation,
    build_text_mapping,
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
            "presentation_builder.extract_xlsx_to_text_mapping",
            return_value={"ROE_RECORRENTE": "10,5%"},
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
            "presentation_builder.generate_chart_assets",
            return_value=11,
        ):
            with patch(
                "presentation_builder.build_text_mapping",
                return_value={"slide1_title": "Titulo"},
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

        self.assertEqual(result.generated_chart_count, 11)
        self.assertEqual(result.replaced_pictures, 3)
        self.assertEqual(result.replaced_text, 2)
        update_mock.assert_called_once()


if __name__ == "__main__":
    unittest.main()
