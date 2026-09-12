from __future__ import annotations

import unittest

from PySide6.QtWidgets import QApplication

from cutmanager.collab import apply_diff, snapshot_rows, diff_snapshots
from cutmanager.constants import COLUMN_CUT_NUMBER, COLUMN_MEMO, CSV_HEADERS
from cutmanager.history import HistoryManager
from cutmanager.model import CutTableModel
from cutmanager.proxy import CutFilterProxyModel
from cutmanager.view import CutTableView


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


class RemoteApplyTest(unittest.TestCase):
    def setUp(self) -> None:
        _app()
        self.model = CutTableModel([_row("001"), _row("002")])
        self.history = HistoryManager(100)
        self.model.set_history_manager(self.history)

    def test_remote_cell_change_does_not_enter_undo_history(self) -> None:
        applied = self.model.apply_remote_cell_changes([(0, COLUMN_MEMO, "他の人の編集")])
        self.assertEqual(applied, 1)
        self.assertEqual(self.model.rows()[0][COLUMN_MEMO], "他の人の編集")
        self.assertFalse(self.history.can_undo())
        self.assertTrue(self.model.is_modified())

    def test_remote_cell_change_emits_data_changed(self) -> None:
        changed: list = []
        self.model.dataChanged.connect(lambda *args: changed.append(args))
        self.model.apply_remote_cell_changes([(1, COLUMN_MEMO, "更新")])
        self.assertEqual(len(changed), 1)

    def test_remote_rows_replace_content_without_history(self) -> None:
        self.model.apply_remote_rows([_row("003", "追加")])
        self.assertEqual([row[COLUMN_CUT_NUMBER] for row in self.model.rows()], ["003"])
        self.assertFalse(self.history.can_undo())

    def test_local_edit_still_uses_history(self) -> None:
        self.model.apply_cell_changes([(0, COLUMN_MEMO, "自分の編集")])
        self.assertTrue(self.history.can_undo())

    def test_diff_then_apply_round_trip_through_the_model(self) -> None:
        before = snapshot_rows(self.model.rows())
        edited = self.model.rows()
        edited[1][COLUMN_MEMO] = "リテイク 2"
        diff = diff_snapshots(before, snapshot_rows(edited))

        _rows, cell_updates, structural = apply_diff(self.model.rows(), diff)
        self.assertFalse(structural)
        self.model.apply_remote_cell_changes(cell_updates)
        self.assertEqual(self.model.rows()[1][COLUMN_MEMO], "リテイク 2")


class RemoteCursorViewTest(unittest.TestCase):
    def setUp(self) -> None:
        _app()
        self.model = CutTableModel([_row("001"), _row("002")])
        self.proxy = CutFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.view = CutTableView()
        self.view.setModel(self.proxy)

    def test_cursors_are_stored_and_repeated_updates_are_ignored(self) -> None:
        cursors = [(1, COLUMN_MEMO, "アリス", "#e5484d")]
        self.view.set_remote_cursors(cursors)
        self.assertEqual(self.view._remote_cursors, cursors)
        # 同じ内容の再設定は再描画を起こさない（値も変わらない）。
        self.view.set_remote_cursors(list(cursors))
        self.assertEqual(self.view._remote_cursors, cursors)

    def test_cursors_can_be_cleared(self) -> None:
        self.view.set_remote_cursors([(0, 0, "アリス", "#0090ff")])
        self.view.set_remote_cursors([])
        self.assertEqual(self.view._remote_cursors, [])

    def test_painting_out_of_range_cursor_is_safe(self) -> None:
        self.view.resize(400, 200)
        self.view.set_remote_cursors([(99, 99, "範囲外", "#30a46c")])
        self.view.viewport().render(self.view.viewport().grab())

    def test_label_text_color_follows_background_brightness(self) -> None:
        from PySide6.QtGui import QColor

        self.assertEqual(CutTableView._label_text_color(QColor("#ffffff")).name(), "#101010")
        self.assertEqual(CutTableView._label_text_color(QColor("#101010")).name(), "#ffffff")


if __name__ == "__main__":
    unittest.main()
