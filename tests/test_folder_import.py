from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cutmanager.folder_import import build_rows_from_dropped_folders


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


if __name__ == "__main__":
    unittest.main()
