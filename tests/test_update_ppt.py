import unittest

from pptx import Presentation

from update_ppt import _replace_text_in_shape


class TestUpdatePpt(unittest.TestCase):
    def _make_text_shape(self, text: str):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_textbox(0, 0, 1000000, 1000000)
        shape.text_frame.text = text
        return shape

    def test_replace_text_in_shape_formats_var_pp_fields_as_pp(self):
        shape = self._make_text_shape("{{VAR_TEST}}")

        replaced = _replace_text_in_shape(
            shape,
            {"VAR_TEST": "-0.9"},
            pp_field_ids={"VAR_TEST"},
        )

        joined = "".join(run.text for paragraph in shape.text_frame.paragraphs for run in paragraph.runs)
        self.assertEqual(replaced, 1)
        self.assertEqual(joined, "▼ 0,9 p.p.")

    def test_replace_text_in_shape_keeps_percent_logic_for_regular_var(self):
        shape = self._make_text_shape("{{VAR_TEST}}")

        replaced = _replace_text_in_shape(
            shape,
            {"VAR_TEST": "-0.9"},
            pp_field_ids=set(),
        )

        joined = "".join(run.text for paragraph in shape.text_frame.paragraphs for run in paragraph.runs)
        self.assertEqual(replaced, 1)
        self.assertEqual(joined, "▼ 90,0%")


if __name__ == "__main__":
    unittest.main()
