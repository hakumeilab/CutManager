"""共同編集が自動で始まる動きと、表示名の扱いを確かめる。

- ファイルを開いた時点で共有に参加し、誰かが同じファイルを開くと共有が始まる
- 相手がいない間は共有フォルダーへ書かず、自動保存もしない
- 相手が現れた時点で、それまでの編集をまとめて配信する
- 表示名は PC のユーザー名ではなく、本人が設定したものだけを使う
"""

from __future__ import annotations

import os
import tempfile
import unittest
from unittest import mock

# 実際のユーザー設定を読み書きしないよう、PySide6 より先に保存先を切り替える。
_settings_dir = tempfile.TemporaryDirectory()
os.environ["XDG_CONFIG_HOME"] = _settings_dir.name
os.environ["APPDATA"] = _settings_dir.name

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication

from cutmanager.collab import CollabSession
from cutmanager.constants import COLUMN_CUT_NUMBER, COLUMN_MEMO, CSV_HEADERS
from cutmanager.csv_io import load_csv_file, save_csv_file
from cutmanager.folder_import import make_cut_key


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


class CollabAutoStartTest(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()
        from cutmanager.main_window import MainWindow

        settings = QSettings("CutManager", "CutManager")
        settings.clear()
        settings.setValue("collab/displayName", "テスト担当")
        settings.sync()

        self._temp_dir = tempfile.TemporaryDirectory()
        self.path = f"{self._temp_dir.name}/共有.cutmgr"
        save_csv_file(self.path, [_row("001"), _row("002")])

        self.window = MainWindow()
        self.window.open_csv_path(self.path)
        self.app.processEvents()
        self.peer: CollabSession | None = None

    def tearDown(self) -> None:
        if self.peer is not None:
            self.peer.stop()
        self.window._stop_collaboration(remember=False)
        self.window._skip_close_confirmation = True
        self.window.close()
        self._temp_dir.cleanup()

    def _join_peer(self, name: str = "ボブ") -> CollabSession:
        self.peer = CollabSession()
        self.peer.start(self.path, name, load_csv_file(self.path).rows)
        self.window.collab.poll()
        self.app.processEvents()
        return self.peer

    def _own_ops_size(self) -> int:
        ops_path = (
            f"{self._temp_dir.name}/共有.cutmgr.cutsync/ops/"
            f"{self.window.collab.session_id}.jsonl"
        )
        return os.path.getsize(ops_path) if os.path.exists(ops_path) else 0

    # --------------------------------------------------------------- 自動開始

    def test_opening_a_file_joins_automatically(self) -> None:
        self.assertTrue(self.window.collab.is_active())
        self.assertEqual(self.window.collab.file_path(), self.path)

    def test_alone_means_waiting(self) -> None:
        self.assertEqual(self.window.collab.peers(), [])
        self.assertIn("自分のみ", self.window.collab_label.text())

    def test_no_shared_writes_while_alone(self) -> None:
        self.window.model.apply_cell_changes([(0, COLUMN_MEMO, "ひとりで編集")])
        self.app.processEvents()
        self.assertFalse(self.window._collab_publish_timer.isActive())
        self.assertEqual(self._own_ops_size(), 0)

    def test_no_autosave_while_alone(self) -> None:
        self.window.model.apply_cell_changes([(0, COLUMN_MEMO, "ひとりで編集")])
        self.app.processEvents()
        self.window._collab_autosave()
        self.assertEqual(load_csv_file(self.path).rows[0][COLUMN_MEMO], "")
        self.assertTrue(self.window.model.is_modified())

    def test_edits_made_alone_reach_the_peer_who_joins_later(self) -> None:
        self.window.model.apply_cell_changes([(0, COLUMN_MEMO, "参加前の編集")])
        self.app.processEvents()

        peer = self._join_peer()
        received: list = []
        peer.remoteDiffReceived.connect(received.append)
        peer.poll()

        self.assertEqual(len(received), 1)
        self.assertEqual(received[0].cells[make_cut_key("001")], {COLUMN_MEMO: "参加前の編集"})

    def test_autosave_runs_while_sharing(self) -> None:
        self._join_peer()
        self.window.model.apply_cell_changes([(1, COLUMN_MEMO, "共有中の編集")])
        self.app.processEvents()
        self.window._collab_autosave()
        self.assertEqual(load_csv_file(self.path).rows[1][COLUMN_MEMO], "共有中の編集")

    def test_peer_arrival_is_announced(self) -> None:
        self._join_peer("演出 佐藤")
        self.assertIn("演出 佐藤", self.window.statusBar().currentMessage())
        self.assertIn("+ 1 人", self.window.collab_label.text())

    # ------------------------------------------------------------ 無効にする

    def test_turning_it_off_stops_sharing_and_is_remembered(self) -> None:
        self.window.collab_action.setChecked(False)
        self.assertFalse(self.window.collab.is_active())
        self.assertFalse(self.window._collaboration_enabled())
        self.assertIn("オフ", self.window.collab_label.text())

    def test_turning_it_back_on_rejoins(self) -> None:
        self.window.collab_action.setChecked(False)
        self.window.collab_action.setChecked(True)
        self.assertTrue(self.window.collab.is_active())
        self.assertTrue(self.window._collaboration_enabled())

    def test_disabled_setting_is_respected_when_the_file_changes(self) -> None:
        self.window.collab_action.setChecked(False)
        other_path = f"{self._temp_dir.name}/別.cutmgr"
        save_csv_file(other_path, [_row("010")])
        self.window.open_csv_path(other_path)
        self.app.processEvents()
        self.assertFalse(self.window.collab.is_active())

    def test_opening_another_file_moves_the_session(self) -> None:
        other_path = f"{self._temp_dir.name}/別.cutmgr"
        save_csv_file(other_path, [_row("010")])
        self.window.open_csv_path(other_path)
        self.app.processEvents()
        self.assertTrue(self.window.collab.is_active())
        self.assertEqual(self.window.collab.file_path(), other_path)

    # -------------------------------------------------------------- 表示名

    def test_display_name_comes_from_the_setting(self) -> None:
        self.assertEqual(self.window._collab_display_name(), "テスト担当")
        self.assertIn("テスト担当", self.window.collab_label.text())

    def test_display_name_does_not_fall_back_to_the_os_user(self) -> None:
        QSettings("CutManager", "CutManager").setValue("collab/displayName", "")
        self.assertEqual(self.window._collab_display_name(), "")
        self.assertEqual(self.window._collab_display_name_or_placeholder(), "名前未設定")

    def test_changing_the_display_name_reaches_the_peer(self) -> None:
        peer = self._join_peer()
        self.window._set_collab_display_name("演出 田中")
        peer.poll()
        self.assertEqual([person.name for person in peer.peers()], ["演出 田中"])

    def test_name_is_asked_once_when_unset(self) -> None:
        settings = QSettings("CutManager", "CutManager")
        settings.setValue("collab/displayName", "")
        settings.setValue("collab/displayNameAsked", False)
        settings.sync()

        with mock.patch(
            "cutmanager.main_window.QInputDialog.getText",
            return_value=("作打ち 鈴木", True),
        ) as prompt:
            self.window._ensure_collab_display_name()
            self.window._ensure_collab_display_name()

        self.assertEqual(prompt.call_count, 1)
        self.assertEqual(self.window._collab_display_name(), "作打ち 鈴木")

    def test_declining_the_prompt_does_not_ask_again(self) -> None:
        settings = QSettings("CutManager", "CutManager")
        settings.setValue("collab/displayName", "")
        settings.setValue("collab/displayNameAsked", False)
        settings.sync()

        with mock.patch(
            "cutmanager.main_window.QInputDialog.getText",
            return_value=("", False),
        ) as prompt:
            self.window._ensure_collab_display_name()
            self.window._ensure_collab_display_name()

        self.assertEqual(prompt.call_count, 1)
        self.assertEqual(self.window._collab_display_name(), "")


if __name__ == "__main__":
    unittest.main()
