"""描画・一括編集を軽くするための最適化が、表示内容や操作結果を変えていないことを確かめる。"""

from __future__ import annotations

import os
import tempfile
import unittest

# 実際のユーザー設定を読み書きしないよう、PySide6 を読み込む前に保存先を切り替える。
_settings_dir = tempfile.TemporaryDirectory()
os.environ["XDG_CONFIG_HOME"] = _settings_dir.name
os.environ["APPDATA"] = _settings_dir.name

from PySide6.QtCore import QItemSelection, QItemSelectionModel, Qt
from PySide6.QtWidgets import QApplication

from cutmanager.constants import (
    BG_STATE_APPROVED,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_BG_STATE,
    COLUMN_CUT_NUMBER,
    COLUMN_MEMO,
    COLUMN_STATUS,
    COLUMN_THUMBNAIL,
    COLUMN_TP_LOAD_COUNT,
    COLUMN_VIDEO_PATH,
    CSV_HEADERS,
)
from cutmanager.history import HistoryManager
from cutmanager.model import CutTableModel
from cutmanager.proxy import CutFilterProxyModel


def _app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _row(cut_number: str, **values: str) -> list[str]:
    row = [""] * len(CSV_HEADERS)
    row[COLUMN_CUT_NUMBER] = cut_number
    for column, value in values.items():
        row[int(column.removeprefix("c"))] = value
    return row


def _model(rows: list[list[str]]) -> tuple[CutTableModel, HistoryManager]:
    model = CutTableModel(rows)
    history = HistoryManager(100)
    model.set_history_manager(history)
    return model, history


BACKGROUND = Qt.ItemDataRole.BackgroundRole
FOREGROUND = Qt.ItemDataRole.ForegroundRole


class RowColorTest(unittest.TestCase):
    def setUp(self) -> None:
        _app()

    def test_plain_row_has_no_colors(self) -> None:
        model, _history = _model([_row("001")])
        self.assertIsNone(model.data(model.index(0, COLUMN_MEMO), BACKGROUND))
        self.assertIsNone(model.data(model.index(0, COLUMN_MEMO), FOREGROUND))

    def test_status_colors_whole_row(self) -> None:
        model, _history = _model([_row("001", **{f"c{COLUMN_STATUS}": "欠番"})])
        backgrounds = {model.data(model.index(0, column), BACKGROUND).name() for column in (0, COLUMN_MEMO, COLUMN_THUMBNAIL)}
        self.assertEqual(len(backgrounds), 1)
        self.assertIsNotNone(model.data(model.index(0, COLUMN_MEMO), FOREGROUND))

    def test_special_cells_override_row_color(self) -> None:
        model, _history = _model(
            [
                _row(
                    "001",
                    **{
                        f"c{COLUMN_STATUS}": "兼用",
                        f"c{COLUMN_TP_LOAD_COUNT}": "BGOnly",
                        f"c{COLUMN_BG_STATE}": BG_STATE_APPROVED,
                    },
                )
            ]
        )
        row_color = model.data(model.index(0, COLUMN_MEMO), BACKGROUND)
        special_color = model.data(model.index(0, COLUMN_TP_LOAD_COUNT), BACKGROUND)
        state_color = model.data(model.index(0, COLUMN_BG_STATE), BACKGROUND)
        self.assertNotEqual(row_color, special_color)
        self.assertNotEqual(row_color, state_color)
        self.assertIsNotNone(model.data(model.index(0, COLUMN_TP_LOAD_COUNT), FOREGROUND))
        # 行の文字色は欠番だけ。兼用の行では特殊セル以外は既定の文字色。
        self.assertIsNone(model.data(model.index(0, COLUMN_MEMO), FOREGROUND))

    def test_colors_follow_edits_and_undo(self) -> None:
        model, history = _model([_row("001"), _row("002")])
        index = model.index(1, COLUMN_MEMO)
        self.assertIsNone(model.data(index, BACKGROUND))

        model.setData(model.index(1, COLUMN_STATUS), "BANK", Qt.ItemDataRole.EditRole)
        self.assertIsNotNone(model.data(index, BACKGROUND))
        model.setData(model.index(1, COLUMN_BG_LOAD_COUNT), "全セル", Qt.ItemDataRole.EditRole)
        self.assertNotEqual(model.data(model.index(1, COLUMN_BG_LOAD_COUNT), BACKGROUND), model.data(index, BACKGROUND))

        history.undo()
        history.undo()
        self.assertIsNone(model.data(index, BACKGROUND))
        self.assertIsNone(model.data(model.index(0, COLUMN_MEMO), BACKGROUND))

    def test_thumbnail_and_virtual_row_values(self) -> None:
        model, _history = _model([_row("001", **{f"c{COLUMN_VIDEO_PATH}": "/tmp/a.mp4"})])
        self.assertEqual(model.data(model.index(0, COLUMN_THUMBNAIL)), "")
        self.assertEqual(model.data(model.index(0, COLUMN_VIDEO_PATH)), "/tmp/a.mp4")
        empty_model, _ = _model([])
        self.assertEqual(empty_model.data(empty_model.index(0, COLUMN_MEMO)), "")
        self.assertIsNone(empty_model.data(empty_model.index(0, COLUMN_MEMO), BACKGROUND))


class BulkChangeTest(unittest.TestCase):
    def setUp(self) -> None:
        _app()

    def test_bulk_change_is_notified_once(self) -> None:
        model, history = _model([_row(f"{index:03d}") for index in range(100)])
        notifications: list = []
        model.dataChanged.connect(lambda top, bottom, roles: notifications.append((top.row(), bottom.row())))

        model.apply_cell_changes([(row, COLUMN_STATUS, "兼用") for row in range(10, 90)])

        self.assertEqual(notifications, [(10, 89)])
        self.assertTrue(all(model.rows()[row][COLUMN_STATUS] == "兼用" for row in range(10, 90)))
        self.assertIsNotNone(model.data(model.index(50, COLUMN_MEMO), BACKGROUND))
        history.undo()
        self.assertTrue(all(row[COLUMN_STATUS] == "" for row in model.rows()))
        self.assertIsNone(model.data(model.index(50, COLUMN_MEMO), BACKGROUND))

    def test_sorting_does_not_corrupt_undo_history(self) -> None:
        model, history = _model([_row("002"), _row("001")])
        model.insert_blank_row(0)
        model.sort(COLUMN_CUT_NUMBER, Qt.SortOrder.DescendingOrder)
        history.undo()
        self.assertEqual([row[COLUMN_CUT_NUMBER] for row in model.rows()], ["002", "001"])


class FilterTest(unittest.TestCase):
    def setUp(self) -> None:
        _app()

    def test_filter_reads_row_values(self) -> None:
        model, _history = _model(
            [_row("001", **{f"c{COLUMN_STATUS}": "兼用"}), _row("002"), _row("003", **{f"c{COLUMN_STATUS}": "BANK"})]
        )
        proxy = CutFilterProxyModel()
        proxy.setSourceModel(model)
        proxy.set_allowed_values(COLUMN_STATUS, {"", "BANK"})
        visible = [proxy.data(proxy.index(row, COLUMN_CUT_NUMBER)) for row in range(proxy.rowCount())]
        self.assertEqual(visible, ["002", "003"])
        proxy.clear_all_filters()
        self.assertEqual(proxy.rowCount(), 3)


class SelectedColorTest(unittest.TestCase):
    def setUp(self) -> None:
        _app()

    def test_selected_colored_cell_is_clearly_distinguishable(self) -> None:
        from PySide6.QtGui import QColor, QPalette

        from cutmanager.view import CutItemDelegate

        background = QColor("#d6f1e1")  # 兼用の行
        selected = CutItemDelegate._selected_fill_color(background, QPalette(QColor("#ffffff")))
        difference = sum(
            abs(a - b)
            for a, b in zip(
                (background.red(), background.green(), background.blue()),
                (selected.red(), selected.green(), selected.blue()),
            )
        )
        # 選択色が淡く混ざるだけで見分けられない状態に戻っていないこと。
        self.assertGreater(difference, 60)
        self.assertGreater(selected.blue(), selected.green() - 40)


class MainWindowEditingTest(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()
        from cutmanager.main_window import MainWindow

        self.window = MainWindow()
        self.window.model.replace_rows([_row(f"{index:03d}") for index in range(10)])
        self.window.history.clear()

    def tearDown(self) -> None:
        self.window._stop_collaboration(remember=False)
        self.window._skip_close_confirmation = True
        self.window.close()

    def _select(self, selection: QItemSelection) -> None:
        self.window.table_view.selectionModel().select(selection, QItemSelectionModel.SelectionFlag.ClearAndSelect)

    def test_selected_rows_and_columns_from_ranges(self) -> None:
        proxy = self.window.proxy_model
        selection = QItemSelection(proxy.index(2, 1), proxy.index(4, 3))
        selection.select(proxy.index(7, 5), proxy.index(7, 5))
        self._select(selection)

        self.assertEqual(self.window._selected_rows(), {2, 3, 4, 7})
        self.assertEqual(self.window._selected_columns(), {1, 2, 3, 5})

    def test_pasting_one_value_into_many_cells_is_one_undo_step(self) -> None:
        proxy = self.window.proxy_model
        self._select(QItemSelection(proxy.index(0, COLUMN_MEMO), proxy.index(5, COLUMN_VIDEO_PATH)))
        QApplication.clipboard().setText("貼り付け")

        self.window.paste_cells_from_clipboard()

        rows = self.window.model.rows()
        self.assertTrue(all(rows[row][COLUMN_MEMO] == "貼り付け" for row in range(6)))
        self.assertEqual(rows[6][COLUMN_MEMO], "")
        # 動画パスはプログラムが管理する列なので貼り付けの対象外。
        self.assertTrue(all(rows[row][COLUMN_VIDEO_PATH] == "" for row in range(6)))

        self.assertTrue(self.window.history.undo())
        self.assertTrue(all(row[COLUMN_MEMO] == "" for row in self.window.model.rows()))
        self.assertFalse(self.window.history.can_undo())

    def test_memo_is_saved_after_typing_settles(self) -> None:
        memo_edit = self.window.summary_memo_edit
        memo_edit.setPlainText("打ち合わせメモ")
        key = self.window._episode_memo_key()
        # 1 文字ごとには書き出さず、少し待ってからまとめて保存する。
        self.assertTrue(self.window._memo_save_timer.isActive())

        self.window._flush_summary_memo()
        self.assertEqual(self.window.settings.value(key), "打ち合わせメモ")
        self.assertEqual(self.window._load_episode_memo(), "打ち合わせメモ")


if __name__ == "__main__":
    unittest.main()
