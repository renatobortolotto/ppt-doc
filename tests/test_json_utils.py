import unittest

from utils.json_utils import coerce_json, first_json_object_slice, strip_fences


class TestJsonUtils(unittest.TestCase):
    def test_strip_fences_removes_code_block(self):
        text = "```json\n{\"a\": 1}\n```"
        self.assertEqual(strip_fences(text), '{"a": 1}')

    def test_first_json_object_slice_extracts_first_object(self):
        s = 'prefix {"a": 1} middle {"b": 2} suffix'
        self.assertEqual(first_json_object_slice(s), '{"a": 1}')

    def test_first_json_object_slice_handles_nested(self):
        s = 'xx {"a": {"b": 2}, "c": 3} yy'
        self.assertEqual(first_json_object_slice(s), '{"a": {"b": 2}, "c": 3}')

    def test_first_json_object_slice_ignores_braces_in_strings(self):
        s = 'xx {"a": "{not a brace}", "b": 1} yy'
        self.assertEqual(first_json_object_slice(s), '{"a": "{not a brace}", "b": 1}')

    def test_coerce_json_parses_plain_json(self):
        self.assertEqual(coerce_json('{"x": 1}'), {"x": 1})

    def test_coerce_json_parses_fenced_json(self):
        self.assertEqual(coerce_json("```\n{\"x\": 1}\n```"), {"x": 1})

    def test_coerce_json_parses_json_embedded_in_text(self):
        out = coerce_json('some text {"x": 1, "y": 2} more')
        self.assertEqual(out, {"x": 1, "y": 2})


if __name__ == "__main__":
    unittest.main()
