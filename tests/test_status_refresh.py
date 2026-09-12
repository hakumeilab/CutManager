"""状態表示の更新をまとめて行う最適化が、表示内容を壊していないことを確かめる。

1 編集ごとに全行を数え直すと行数の多いファイルで重くなるため、
状態表示はイベントループ 1 周に 1 回、進捗リボンはさらに間引いて更新する。
"""

from __future__ import annotations

import os
import tempfile
import unittest

# 実際のユーザー設定（前回開いたファイルや共同編集の設定）を読み書きしないよう、
# PySide6 を読み込む前に設定の保存先を切り替える。
_settings_dir = tempfile.TemporaryDirectory()
os.environ["XDG_CONFIG_HOME"] = _settings_dir.name
os.environ["APPDATA"] = _settings_dir.name

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from cutmanager.constants import (
    COLUMN_CUT_NUMBER,
    COLUMN_DELIVERY_DATE,
    COLUMN_MEMO,
    COLUMN_STATUS,
    CSV_HEADERS,
)
from cutmanager.csv_io import save_csv_file


def _app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _row(cut_number: str, delivered: str = "", status: str = "") -> list[str]:
    row = [""] * len(CSV_HEADERS)
    row[COLUMN_CUT_NUMBER] = cut_number
    row[COLUMN_DELIVERY_DATE] = delivered
    row[COLUMN_STATUS] = status
    return row


class StatusRefreshTest(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()
        from cutmanager.main_window import MainWindow

        self._temp_dir = tempfile.TemporaryDirectory()
        self.path = f"{self._temp_dir.name}/status.cutmgr"
        save_csv_file(
            self.path,
            [_row("001", "2026/01/01"), _row("002"), _row("003", status="欠番")],
        )
        # 共同編集は既定で有効。表示名を先に入れて、初回の名前入力を出さない。
        settings = QSettings("CutManager", "CutManager")
        settings.setValue("collab/displayName", "テスト担当")
        settings.sync()

        self.window = MainWindow()
        self.window.open_csv_path(self.path)
        self.app.processEvents()

    def tearDown(self) -> None:
        self.window._stop_collaboration(remember=False)
        # 未保存の確認ダイアログで止まらないように閉じる。
        self.window._skip_close_confirmation = True
        self.window.close()
        self._temp_dir.cleanup()

    def _summary_text(self, key: str) -> str:
        return self.window.summary_labels[key].text()

    def test_summary_is_correct_after_load(self) -> None:
        self.window._update_summary_ribbon()
        self.assertTrue(self._summary_text("total_cuts").endswith("2"))
        self.assertTrue(self._summary_text("delivered").endswith("1"))
        self.assertTrue(self._summary_text("missing").endswith("1"))

    def test_edit_schedules_a_summary_refresh(self) -> None:
        self.window._update_summary_ribbon()
        self.window.model.apply_cell_changes([(1, COLUMN_DELIVERY_DATE, "2026/02/02")])
        self.app.processEvents()
        self.assertTrue(self.window._summary_refresh_timer.isActive())

        # 間引きのタイマーが動いたあとは最新の数字になる。
        self.window._update_summary_ribbon()
        self.assertTrue(self._summary_text("delivered").endswith("2"))

    def test_status_labels_update_after_the_event_loop_turn(self) -> None:
        self.window.model.apply_cell_changes([(0, COLUMN_MEMO, "編集")])
        self.app.processEvents()
        self.assertIn("未保存", self.window.modified_label.text())
        self.assertIn("行数: 3", self.window.row_count_label.text())

    def test_repeated_edits_do_not_pile_up_status_updates(self) -> None:
        calls = [0]
        original = self.window._update_status_labels

        def counted() -> None:
            calls[0] += 1
            original()

        self.window._update_status_labels = counted
        for index in range(10):
            self.window.model.apply_cell_changes([(0, COLUMN_MEMO, f"編集{index}")])
        self.app.processEvents()
        self.assertEqual(calls[0], 1)

    def test_row_count_follows_row_insertion(self) -> None:
        self.window.model.insert_blank_row(0)
        self.app.processEvents()
        self.assertIn("行数: 4", self.window.row_count_label.text())


if __name__ == "__main__":
    unittest.main()
