from __future__ import annotations

import unittest

from PySide6.QtCore import QItemSelectionModel, Qt
from PySide6.QtWidgets import QApplication, QTableView

from cutmanager.constants import COLUMN_CUT_NUMBER, COLUMN_MEMO
from cutmanager.history import HistoryManager
from cutmanager.model import CutTableModel
from cutmanager.proxy import CutFilterProxyModel
from cutmanager.view import CellEditorLineEdit, CutItemDelegate


def _app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _make_stack():
    model = CutTableModel(
        [
            ["001", "", "", "", "", "", "", "", "", "", ""],
            ["002", "", "", "", "", "", "", "", "", "", ""],
            ["003", "", "", "", "", "", "", "", "", "", ""],
        ]
    )
    history = HistoryManager(100)
    model.set_history_manager(history)
    proxy = CutFilterProxyModel()
    proxy.setSourceModel(model)
    view = QTableView()
    view.setModel(proxy)
    delegate = CutItemDelegate(view)
    view.setItemDelegate(delegate)
    return model, history, proxy, view, delegate


class MultiCellFillTests(unittest.TestCase):
    def setUp(self) -> None:
        _app()

    def test_edit_fills_all_selected_cells_in_column(self) -> None:
        model, history, proxy, view, delegate = _make_stack()
        selection = view.selectionModel()
        for row in (0, 1, 2):
            selection.select(
                proxy.index(row, COLUMN_MEMO),
                QItemSelectionModel.SelectionFlag.Select,
            )

        edited = proxy.index(0, COLUMN_MEMO)
        editor = CellEditorLineEdit()
        editor.setText("一括メモ")
        delegate.setModelData(editor, proxy, edited)

        for row in (0, 1, 2):
            self.assertEqual(model.rows()[row][COLUMN_MEMO], "一括メモ")

        # まとめて 1 回のアンドゥで元に戻せること。
        self.assertTrue(history.undo())
        for row in (0, 1, 2):
            self.assertEqual(model.rows()[row][COLUMN_MEMO], "")

    def test_single_selection_only_edits_one_cell(self) -> None:
        model, history, proxy, view, delegate = _make_stack()
        view.selectionModel().select(
            proxy.index(1, COLUMN_MEMO),
            QItemSelectionModel.SelectionFlag.Select,
        )
        editor = CellEditorLineEdit()
        editor.setText("単独")
        delegate.setModelData(editor, proxy, proxy.index(1, COLUMN_MEMO))

        self.assertEqual(model.rows()[1][COLUMN_MEMO], "単独")
        self.assertEqual(model.rows()[0][COLUMN_MEMO], "")
        self.assertEqual(model.rows()[2][COLUMN_MEMO], "")


class StickyFilterTests(unittest.TestCase):
    def setUp(self) -> None:
        _app()

    def test_edited_row_stays_visible_until_filter_reapplied(self) -> None:
        model = CutTableModel(
            [
                ["001", "A", "", "", "", "", "", "", "", "", ""],
                ["002", "B", "", "", "", "", "", "", "", "", ""],
            ]
        )
        model.set_history_manager(HistoryManager(100))
        proxy = CutFilterProxyModel()
        proxy.setSourceModel(model)
        proxy.set_allowed_values(1, {"A"})  # AB分け列で "A" のみ表示
        self.assertEqual(proxy.rowCount(), 1)

        # フィルター対象の値を編集しても、行はすぐには絞り込みから外れない。
        source_index = model.index(0, 1)
        model.setData(source_index, "B", Qt.ItemDataRole.EditRole)
        self.assertEqual(proxy.rowCount(), 1)

        # フィルターを掛け直すと再評価され、行が外れる。
        proxy.set_allowed_values(1, {"A"})
        self.assertEqual(proxy.rowCount(), 0)


if __name__ == "__main__":
    unittest.main()
