import importlib
import unittest

import src.utils as utils


class TestSlidePackageStructure(unittest.TestCase):
    def test_slide_modules_live_under_new_package(self):
        module = importlib.import_module("src.utils.slides.slide4_charts")

        self.assertTrue(hasattr(module, "generate_slide4_charts"))

    def test_old_slide_module_path_is_not_available_anymore(self):
        with self.assertRaises(ModuleNotFoundError):
            importlib.import_module("utils.slide4_charts")

    def test_utils_root_no_longer_exports_slide_generators(self):
        self.assertFalse(hasattr(utils, "generate_slide4_charts"))
        self.assertFalse(hasattr(utils, "generate_pizza_charts"))


if __name__ == "__main__":
    unittest.main()
