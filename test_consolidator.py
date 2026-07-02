from pathlib import Path
import unittest

from dcsc_mis_consolidator import (
    airline_from_path,
    direction_from_path,
    fortnight_from_text,
    month_from_text,
    split_awb_suffix,
)


class ConsolidatorTests(unittest.TestCase):
    def test_metadata_from_path_and_filename(self):
        root = Path("pack")
        path = root / "Export" / "AY" / "AY EXPORT FEB'2026 (01-15).xlsx"

        self.assertEqual(direction_from_path(root, path), "EXP")
        self.assertEqual(airline_from_path(root, path), "AY")
        self.assertEqual(month_from_text(path.name), "FEB'2026")
        self.assertEqual(fortnight_from_text(path.name), "01-15")

    def test_combined_awb_suffix_split(self):
        self.assertEqual(split_awb_suffix("784-27933058A"), ("784-27933058", "A"))
        self.assertEqual(split_awb_suffix("771-93514179", "P"), ("771-93514179", "P"))


if __name__ == "__main__":
    unittest.main()
