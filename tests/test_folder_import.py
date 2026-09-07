from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cutmanager.constants import (
    BG_STATE_APPROVED,
    BG_STATE_RAW,
    COLUMN_AB_GROUP,
    COLUMN_BG_DATE,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_CUT_NUMBER,
    COLUMN_BG_STATE,
    COLUMN_TP_DATE,
    COLUMN_TP_LOAD_COUNT,
    COLUMN_TP_STATE,
    TP_STATE_CHECKED,
    TP_STATE_UNCHECKED,
)
from cutmanager.folder_import import (
    apply_material_updates,
    build_rows_from_dropped_folders,
    extract_cut_identifiers,
)


class MaterialLoadCountTests(unittest.TestCase):
    def test_direct_final_material_starts_at_one(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "materials"
            root.mkdir()
            (root / "001").mkdir()

            # 仮素材を経ずに本番素材が入るケース（状態が入力済みの既存行）。
            rows = [["001", "", "", "", "", "", TP_STATE_CHECKED, "", "", "", "", "", ""]]
            result = build_rows_from_dropped_folders([root], {("001", "")}, "2026/04/16")
            rows = apply_material_updates(rows, result.updates)

        self.assertEqual(rows[0][COLUMN_TP_LOAD_COUNT], "1")
        self.assertEqual(rows[0][COLUMN_TP_STATE], TP_STATE_CHECKED)

    def test_first_material_is_not_counted_and_retakes_are(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "materials"
            root.mkdir()
            (root / "001").mkdir()

            first = build_rows_from_dropped_folders([root], set(), "2026/04/16")
            existing_keys = {("001", "")}
            rows = first.rows
            second = build_rows_from_dropped_folders([root], existing_keys, "2026/04/17")
            rows = apply_material_updates(rows, second.updates)
            third = build_rows_from_dropped_folders([root], existing_keys, "2026/04/18")
            rows = apply_material_updates(rows, third.updates)

        self.assertEqual(first.rows[0][COLUMN_TP_LOAD_COUNT], "0")
        self.assertEqual(first.rows[0][COLUMN_TP_STATE], TP_STATE_UNCHECKED)
        # 2 回目 = 本番素材で 1、3 回目 = リテイクで 2。
        self.assertEqual(rows[0][COLUMN_TP_LOAD_COUNT], "2")
        self.assertEqual(rows[0][COLUMN_TP_DATE], "2026/04/18")


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
        # フォルダー素材は TP 扱いなので未検査から始まる。
        self.assertEqual(result.rows[0][COLUMN_TP_STATE], TP_STATE_UNCHECKED)
        self.assertEqual(result.rows[0][COLUMN_BG_STATE], "")

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
        # 1 回目の素材入れはリテイクに数えないので 0。
        self.assertEqual(result.rows[0][COLUMN_BG_LOAD_COUNT], "0")
        self.assertEqual(result.rows[0][COLUMN_BG_DATE], "2026/04/16")
        self.assertEqual(result.rows[0][COLUMN_BG_STATE], BG_STATE_RAW)
        self.assertEqual(result.rows[0][COLUMN_TP_STATE], "")

    def test_psd_file_updates_existing_bg_count(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            bg_file = Path(temp_dir) / "345_bg.psd"
            bg_file.write_bytes(b"psd")

            result = build_rows_from_dropped_folders([bg_file], {("345", "")}, "2026/04/16")
            rows = [
                ["345", "", "", "", "1", "2026/04/15", "", "2", "2026/04/15", BG_STATE_APPROVED, "", "", ""]
            ]  # BG は既に 2 回入っており、次の素材入れで 3 回目になる
            updated_rows = apply_material_updates(rows, result.updates)

        self.assertEqual(result.added_count, 0)
        self.assertEqual(result.updated_count, 1)
        self.assertEqual(updated_rows[0][COLUMN_TP_LOAD_COUNT], "1")
        self.assertEqual(updated_rows[0][COLUMN_TP_DATE], "2026/04/15")
        self.assertEqual(updated_rows[0][COLUMN_BG_LOAD_COUNT], "3")
        self.assertEqual(updated_rows[0][COLUMN_BG_DATE], "2026/04/16")
        # 入力済みの BG 状態は取り込みで上書きしない。
        self.assertEqual(updated_rows[0][COLUMN_BG_STATE], BG_STATE_APPROVED)


if __name__ == "__main__":
    unittest.main()
