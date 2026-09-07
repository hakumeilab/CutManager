from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cutmanager.constants import (
    COLUMN_AB_GROUP,
    COLUMN_BG_DATE,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_CUT_NUMBER,
    COLUMN_TP_DATE,
    COLUMN_TP_LOAD_COUNT,
)
from cutmanager.folder_import import (
    apply_material_updates,
    build_rows_from_dropped_folders,
    extract_cut_identifiers,
)


class CutNumberExtractionTests(unittest.TestCase):
    def _identifiers(self, name: str) -> list[tuple[str, str]]:
        return [(item.cut_number, item.ab_group) for item in extract_cut_identifiers(name)]

    def test_two_digit_cut_number(self) -> None:
        self.assertEqual(self._identifiers("12"), [("12", "")])
        self.assertEqual(self._identifiers("85A"), [("85", "A")])
        self.assertEqual(self._identifiers("12_13"), [("12", ""), ("13", "")])

    def test_four_digit_cut_number(self) -> None:
        self.assertEqual(self._identifiers("1024"), [("1024", "")])
        self.assertEqual(self._identifiers("cut1024B"), [("1024", "B")])

    def test_three_digit_cut_number_is_unchanged(self) -> None:
        self.assertEqual(self._identifiers("169_170"), [("169", ""), ("170", "")])
        self.assertEqual(self._identifiers("085A"), [("085", "A")])

    def test_roll_and_date_are_not_treated_as_cut_number(self) -> None:
        self.assertEqual(self._identifiers("BMUM_roll01_260721"), [])
        self.assertEqual(self._identifiers("take02"), [])

    def test_five_or_more_digits_are_ignored(self) -> None:
        self.assertEqual(self._identifiers("12345"), [])
        self.assertEqual(self._identifiers("1"), [])


class FolderImportTests(unittest.TestCase):
    def test_parent_folder_imports_named_child_folders(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "materials"
            root.mkdir()
            (root / "001").mkdir()
            (root / "002A").mkdir()
            (root / "12").mkdir()
            (root / "1024B").mkdir()

            result = build_rows_from_dropped_folders([root], set(), "2026/04/16")

        self.assertEqual(result.added_count, 4)
        self.assertEqual(result.failed_count, 0)
        self.assertEqual(
            {(row[COLUMN_CUT_NUMBER], row[COLUMN_AB_GROUP]) for row in result.rows},
            {("001", ""), ("002", "A"), ("12", ""), ("1024", "B")},
        )

    def test_single_cut_folder_is_imported_even_if_it_contains_children(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            cut_folder = Path(temp_dir) / "123A"
            cut_folder.mkdir()
            (cut_folder / "frames").mkdir()
            (cut_folder / "proxy").mkdir()

            result = build_rows_from_dropped_folders([cut_folder], set(), "2026/04/16")

        self.assertEqual(result.added_count, 1)
        self.assertEqual(result.failed_count, 0)
        self.assertEqual(result.rows[0][COLUMN_CUT_NUMBER], "123")
        self.assertEqual(result.rows[0][COLUMN_AB_GROUP], "A")

    def test_psd_file_is_imported_as_bg(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            bg_file = Path(temp_dir) / "345_bg.psd"
            bg_file.write_bytes(b"psd")

            result = build_rows_from_dropped_folders([bg_file], set(), "2026/04/16")

        self.assertEqual(result.added_count, 1)
        self.assertEqual(result.failed_count, 0)
        self.assertEqual(result.rows[0][COLUMN_CUT_NUMBER], "345")
        self.assertEqual(result.rows[0][COLUMN_TP_LOAD_COUNT], "")
        self.assertEqual(result.rows[0][COLUMN_TP_DATE], "")
        self.assertEqual(result.rows[0][COLUMN_BG_LOAD_COUNT], "1")
        self.assertEqual(result.rows[0][COLUMN_BG_DATE], "2026/04/16")

    def test_psd_file_updates_existing_bg_count(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            bg_file = Path(temp_dir) / "345_bg.psd"
            bg_file.write_bytes(b"psd")

            result = build_rows_from_dropped_folders([bg_file], {("345", "")}, "2026/04/16")
            rows = [["345", "", "", "", "1", "2026/04/15", "2", "2026/04/15", "", "", ""]]
            updated_rows = apply_material_updates(rows, result.updates)

        self.assertEqual(result.added_count, 0)
        self.assertEqual(result.updated_count, 1)
        self.assertEqual(updated_rows[0][COLUMN_TP_LOAD_COUNT], "1")
        self.assertEqual(updated_rows[0][COLUMN_TP_DATE], "2026/04/15")
        self.assertEqual(updated_rows[0][COLUMN_BG_LOAD_COUNT], "3")
        self.assertEqual(updated_rows[0][COLUMN_BG_DATE], "2026/04/16")


if __name__ == "__main__":
    unittest.main()
