from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cutmanager.folder_import import apply_material_updates, build_rows_from_dropped_folders


class FolderImportTests(unittest.TestCase):
    def test_parent_folder_imports_named_child_folders(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "materials"
            root.mkdir()
            (root / "001").mkdir()
            (root / "002A").mkdir()

            result = build_rows_from_dropped_folders([root], set(), "2026/04/16")

        self.assertEqual(result.added_count, 2)
        self.assertEqual(result.failed_count, 0)
        self.assertEqual({(row[0], row[1]) for row in result.rows}, {("001", ""), ("002", "A")})

    def test_single_cut_folder_is_imported_even_if_it_contains_children(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            cut_folder = Path(temp_dir) / "123A"
            cut_folder.mkdir()
            (cut_folder / "frames").mkdir()
            (cut_folder / "proxy").mkdir()

            result = build_rows_from_dropped_folders([cut_folder], set(), "2026/04/16")

        self.assertEqual(result.added_count, 1)
        self.assertEqual(result.failed_count, 0)
        self.assertEqual(result.rows[0][0], "123")
        self.assertEqual(result.rows[0][1], "A")

    def test_psd_file_is_imported_as_bg(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            bg_file = Path(temp_dir) / "345_bg.psd"
            bg_file.write_bytes(b"psd")

            result = build_rows_from_dropped_folders([bg_file], set(), "2026/04/16")

        self.assertEqual(result.added_count, 1)
        self.assertEqual(result.failed_count, 0)
        self.assertEqual(result.rows[0][0], "345")
        self.assertEqual(result.rows[0][3], "")
        self.assertEqual(result.rows[0][4], "")
        self.assertEqual(result.rows[0][5], "1")
        self.assertEqual(result.rows[0][6], "2026/04/16")

    def test_psd_file_updates_existing_bg_count(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            bg_file = Path(temp_dir) / "345_bg.psd"
            bg_file.write_bytes(b"psd")

            result = build_rows_from_dropped_folders([bg_file], {("345", "")}, "2026/04/16")
            rows = [["345", "", "", "1", "2026/04/15", "2", "2026/04/15", "", "", ""]]
            updated_rows = apply_material_updates(rows, result.updates)

        self.assertEqual(result.added_count, 0)
        self.assertEqual(result.updated_count, 1)
        self.assertEqual(updated_rows[0][3], "1")
        self.assertEqual(updated_rows[0][4], "2026/04/15")
        self.assertEqual(updated_rows[0][5], "3")
        self.assertEqual(updated_rows[0][6], "2026/04/16")


if __name__ == "__main__":
    unittest.main()
