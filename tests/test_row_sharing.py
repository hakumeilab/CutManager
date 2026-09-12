"""行を共有して持つ最適化が、履歴やスナップショットを壊さないことを確かめる。

行は「書き換えず差し替える」運用にして、アンドゥ履歴やモデルの間で同じ行を
共有している。共有が壊れると、過去の履歴が後からの編集で書き換わってしまう。
"""

from __future__ import annotations

import unittest

from PySide6.QtWidgets import QApplication

from cutmanager.constants import COLUMN_CUT_NUMBER, COLUMN_MEMO, COLUMN_STATUS, CSV_HEADERS
from cutmanager.history import HistoryManager
from cutmanager.model import CutTableModel


def _app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _row(cut_number: str, memo: str = "") -> list[str]:
    row = [""] * len(CSV_HEADERS)
    row[COLUMN_CUT_NUMBER] = cut_number
    row[COLUMN_MEMO] = memo
    return row


class RowSharingTest(unittest.TestCase):
    def setUp(self) -> None:
        _app()
        self.model = CutTableModel([_row("001", "あ"), _row("002", "い"), _row("003", "う")])
        self.history = HistoryManager(100)
        self.model.set_history_manager(self.history)

    def test_edit_then_undo_restores_value(self) -> None:
        self.model.apply_cell_changes([(0, COLUMN_MEMO, "編集後")])
        self.assertEqual(self.model.rows()[0][COLUMN_MEMO], "編集後")
        self.history.undo()
        self.assertEqual(self.model.rows()[0][COLUMN_MEMO], "あ")
        self.history.redo()
        self.assertEqual(self.model.rows()[0][COLUMN_MEMO], "編集後")

    def test_edit_after_row_insert_does_not_corrupt_history(self) -> None:
        self.model.insert_blank_row(0)
        self.model.apply_cell_changes([(1, COLUMN_MEMO, "後から編集")])
        self.assertEqual(self.model.rows()[1][COLUMN_MEMO], "後から編集")

        self.history.undo()  # セル編集を戻す
        self.assertEqual(self.model.rows()[1][COLUMN_MEMO], "あ")
        self.history.undo()  # 行追加を戻す
        self.assertEqual(
            [row[COLUMN_MEMO] for row in self.model.rows()],
            ["あ", "い", "う"],
        )

    def test_repeated_edits_to_one_cell_undo_one_by_one(self) -> None:
        for value in ("1", "2", "3"):
            self.model.apply_cell_changes([(0, COLUMN_MEMO, value)])
        for expected in ("2", "1", "あ"):
            self.history.undo()
            self.assertEqual(self.model.rows()[0][COLUMN_MEMO], expected)

    def test_rows_returns_independent_copies(self) -> None:
        snapshot = self.model.rows()
        snapshot[0][COLUMN_MEMO] = "外から書き換え"
        self.assertEqual(self.model.rows()[0][COLUMN_MEMO], "あ")

    def test_caller_rows_are_not_aliased_into_the_model(self) -> None:
        source = [_row("010", "元")]
        self.model.replace_rows(source)
        self.model.apply_cell_changes([(0, COLUMN_MEMO, "モデル側で編集")])
        self.assertEqual(source[0][COLUMN_MEMO], "元")

    def test_delete_then_undo_keeps_row_content(self) -> None:
        self.model.apply_cell_changes([(1, COLUMN_MEMO, "消す前")])
        self.model.remove_rows_by_numbers([1])
        self.assertEqual([row[COLUMN_CUT_NUMBER] for row in self.model.rows()], ["001", "003"])
        self.history.undo()
        self.assertEqual(self.model.rows()[1][COLUMN_MEMO], "消す前")

    def test_status_edit_after_undo_does_not_leak_into_history(self) -> None:
        self.model.apply_cell_changes([(0, COLUMN_STATUS, "兼用")])
        self.history.undo()
        self.model.apply_cell_changes([(0, COLUMN_STATUS, "BANK")])
        self.history.undo()
        self.assertEqual(self.model.rows()[0][COLUMN_STATUS], "")

    def test_remote_changes_also_copy_before_writing(self) -> None:
        self.model.apply_cell_changes([(0, COLUMN_MEMO, "自分")])
        self.model.apply_remote_cell_changes([(0, COLUMN_MEMO, "相手")])
        self.assertEqual(self.model.rows()[0][COLUMN_MEMO], "相手")
        self.history.undo()
        self.assertEqual(self.model.rows()[0][COLUMN_MEMO], "あ")

    def test_normalized_rows_are_reused_without_copying(self) -> None:
        source = _row("020")
        self.assertIs(CutTableModel._normalize_row(source), source)

    def test_short_rows_are_padded(self) -> None:
        padded = CutTableModel._normalize_row(["001", "メモ"])
        self.assertEqual(len(padded), len(CSV_HEADERS))
        self.assertEqual(padded[COLUMN_MEMO], "メモ")

    def test_non_string_cells_are_converted(self) -> None:
        normalized = CutTableModel._normalize_row([1, None] + [""] * (len(CSV_HEADERS) - 2))
        self.assertEqual(normalized[0], "1")
        self.assertEqual(normalized[1], "")


if __name__ == "__main__":
    unittest.main()
