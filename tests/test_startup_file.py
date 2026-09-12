from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from main import resolve_startup_file


class ResolveStartupFileTest(unittest.TestCase):
    def setUp(self) -> None:
        self._temp_dir = tempfile.TemporaryDirectory()
        self.base = Path(self._temp_dir.name)

    def tearDown(self) -> None:
        self._temp_dir.cleanup()

    def _touch(self, name: str) -> str:
        path = self.base / name
        path.write_text("カット番号\n", encoding="utf-8-sig")
        return str(path)

    def test_no_arguments(self) -> None:
        self.assertIsNone(resolve_startup_file(["CutManager.exe"]))

    def test_project_file_is_opened(self) -> None:
        path = self._touch("cut_list.cutmgr")
        self.assertEqual(resolve_startup_file(["CutManager.exe", path]), path)

    def test_csv_is_accepted(self) -> None:
        path = self._touch("cut_list.csv")
        self.assertEqual(resolve_startup_file(["CutManager.exe", path]), path)

    def test_extension_is_case_insensitive(self) -> None:
        path = self._touch("cut_list.CUTMGR")
        self.assertEqual(resolve_startup_file(["CutManager.exe", path]), path)

    def test_unsupported_extension_is_ignored(self) -> None:
        path = self._touch("memo.txt")
        self.assertIsNone(resolve_startup_file(["CutManager.exe", path]))

    def test_missing_file_is_ignored(self) -> None:
        self.assertIsNone(
            resolve_startup_file(["CutManager.exe", str(self.base / "存在しない.cutmgr")])
        )

    def test_directory_is_ignored(self) -> None:
        folder = self.base / "cut_list.cutmgr"
        folder.mkdir()
        self.assertIsNone(resolve_startup_file(["CutManager.exe", str(folder)]))

    def test_option_like_arguments_are_skipped(self) -> None:
        path = self._touch("cut_list.cutmgr")
        self.assertEqual(resolve_startup_file(["CutManager.exe", "--debug", path]), path)

    def test_first_supported_file_wins(self) -> None:
        first = self._touch("a.cutmgr")
        self._touch("b.cutmgr")
        self.assertEqual(
            resolve_startup_file(["CutManager.exe", first, str(self.base / "b.cutmgr")]),
            first,
        )


if __name__ == "__main__":
    unittest.main()
