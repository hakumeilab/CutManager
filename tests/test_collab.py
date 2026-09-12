from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from PySide6.QtCore import QCoreApplication

from cutmanager.collab import (
    CollabSession,
    RowDiff,
    apply_diff,
    diff_snapshots,
    diff_to_payload,
    payload_to_diff,
    peer_color,
    row_key,
    snapshot_rows,
    sync_dir_for,
)
from cutmanager.constants import (
    COLUMN_AB_GROUP,
    COLUMN_CUT_NUMBER,
    COLUMN_MEMO,
    COLUMN_THUMBNAIL,
    CSV_HEADERS,
)


def _app() -> QCoreApplication:
    app = QCoreApplication.instance()
    if app is None:
        app = QCoreApplication([])
    return app


def _row(cut_number: str, memo: str = "", ab_group: str = "") -> list[str]:
    row = [""] * len(CSV_HEADERS)
    row[COLUMN_CUT_NUMBER] = cut_number
    row[COLUMN_MEMO] = memo
    row[COLUMN_AB_GROUP] = ab_group
    return row


class RowKeyTest(unittest.TestCase):
    def test_key_uses_cut_number_and_ab_group(self) -> None:
        self.assertEqual(row_key(_row("001", ab_group="a")), ("001", "A"))

    def test_row_without_cut_number_has_no_key(self) -> None:
        self.assertIsNone(row_key(_row("", memo="下書き")))
        self.assertIsNone(row_key([]))

    def test_snapshot_keeps_first_row_for_duplicate_keys(self) -> None:
        snapshot = snapshot_rows([_row("001", "先"), _row("001", "後"), _row("", "無視")])
        self.assertEqual(list(snapshot), [("001", "")])
        self.assertEqual(snapshot[("001", "")][COLUMN_MEMO], "先")


class DiffTest(unittest.TestCase):
    def test_detects_changed_cells(self) -> None:
        old = snapshot_rows([_row("001", "旧")])
        new = snapshot_rows([_row("001", "新")])
        diff = diff_snapshots(old, new)
        self.assertEqual(diff.cells, {("001", ""): {COLUMN_MEMO: "新"}})
        self.assertFalse(diff.added)
        self.assertFalse(diff.removed)

    def test_detects_added_and_removed_rows(self) -> None:
        diff = diff_snapshots(snapshot_rows([_row("001")]), snapshot_rows([_row("002")]))
        self.assertEqual(list(diff.added), [("002", "")])
        self.assertEqual(diff.removed, (("001", ""),))

    def test_thumbnail_column_is_not_shared(self) -> None:
        old_row = _row("001")
        new_row = _row("001")
        new_row[COLUMN_THUMBNAIL] = "C:/local/cache/001.png"
        diff = diff_snapshots(snapshot_rows([old_row]), snapshot_rows([new_row]))
        self.assertTrue(diff.is_empty())

    def test_identical_snapshots_produce_empty_diff(self) -> None:
        snapshot = snapshot_rows([_row("001", "メモ")])
        self.assertTrue(diff_snapshots(snapshot, snapshot).is_empty())


class ApplyDiffTest(unittest.TestCase):
    def test_cell_update_keeps_row_order_and_reports_cells(self) -> None:
        rows = [_row("001", "旧"), _row("002")]
        diff = RowDiff(cells={("001", ""): {COLUMN_MEMO: "新"}}, added={}, removed=())
        new_rows, cell_updates, structural = apply_diff(rows, diff)
        self.assertFalse(structural)
        self.assertEqual(cell_updates, [(0, COLUMN_MEMO, "新")])
        self.assertEqual(new_rows[0][COLUMN_MEMO], "新")

    def test_unknown_key_is_ignored_for_cell_updates(self) -> None:
        rows = [_row("001")]
        diff = RowDiff(cells={("999", ""): {COLUMN_MEMO: "x"}}, added={}, removed=())
        new_rows, cell_updates, structural = apply_diff(rows, diff)
        self.assertEqual(new_rows, rows)
        self.assertEqual(cell_updates, [])
        self.assertFalse(structural)

    def test_added_row_is_appended_as_structural_change(self) -> None:
        rows = [_row("001")]
        diff = RowDiff(cells={}, added={("002", ""): _row("002", "追加")}, removed=())
        new_rows, cell_updates, structural = apply_diff(rows, diff)
        self.assertTrue(structural)
        self.assertEqual(cell_updates, [])
        self.assertEqual([row[COLUMN_CUT_NUMBER] for row in new_rows], ["001", "002"])

    def test_removed_row_is_dropped(self) -> None:
        rows = [_row("001"), _row("002")]
        diff = RowDiff(cells={}, added={}, removed=(("001", ""),))
        new_rows, _cell_updates, structural = apply_diff(rows, diff)
        self.assertTrue(structural)
        self.assertEqual([row[COLUMN_CUT_NUMBER] for row in new_rows], ["002"])

    def test_existing_row_is_not_duplicated_by_added_entry(self) -> None:
        rows = [_row("001", "こちらの値")]
        diff = RowDiff(cells={}, added={("001", ""): _row("001", "あちらの値")}, removed=())
        new_rows, _cell_updates, structural = apply_diff(rows, diff)
        self.assertEqual(len(new_rows), 1)
        self.assertFalse(structural)

    def test_cut_number_change_is_treated_as_structural(self) -> None:
        rows = [_row("001")]
        diff = RowDiff(cells={("001", ""): {COLUMN_CUT_NUMBER: "001A"}}, added={}, removed=())
        new_rows, cell_updates, structural = apply_diff(rows, diff)
        self.assertTrue(structural)
        self.assertEqual(cell_updates, [])
        self.assertEqual(new_rows[0][COLUMN_CUT_NUMBER], "001A")


class PayloadTest(unittest.TestCase):
    def test_round_trip(self) -> None:
        diff = RowDiff(
            cells={("001", "A"): {COLUMN_MEMO: "メモ"}},
            added={("002", ""): _row("002")},
            removed=(("003", ""),),
        )
        restored = payload_to_diff(diff_to_payload(diff))
        self.assertEqual(restored.cells, diff.cells)
        self.assertEqual(restored.added, diff.added)
        self.assertEqual(restored.removed, diff.removed)

    def test_broken_entries_are_skipped(self) -> None:
        restored = payload_to_diff({"cells": ["壊れた"], "added": [None], "removed": [["001"]]})
        self.assertTrue(restored.is_empty())

    def test_peer_color_is_stable(self) -> None:
        self.assertEqual(peer_color("abc123"), peer_color("abc123"))
        self.assertTrue(peer_color("abc123").startswith("#"))


class SyncDirTest(unittest.TestCase):
    def test_sidecar_sits_next_to_the_project_file(self) -> None:
        path = sync_dir_for("/mnt/nas/作品/cut_list.cutmgr")
        self.assertEqual(path.name, "cut_list.cutmgr.cutsync")
        self.assertEqual(path.parent, Path("/mnt/nas/作品"))


class CollabSessionTest(unittest.TestCase):
    """2 セッションをサイドカー経由でつないで、実際の往復を確認する。"""

    def setUp(self) -> None:
        _app()
        self._temp_dir = tempfile.TemporaryDirectory()
        self.file_path = str(Path(self._temp_dir.name) / "cut_list.cutmgr")
        self.alice = CollabSession()
        self.bob = CollabSession()

    def tearDown(self) -> None:
        self.alice.stop()
        self.bob.stop()
        self._temp_dir.cleanup()

    def _received(self, session: CollabSession) -> list:
        received: list = []
        session.remoteDiffReceived.connect(received.append)
        return received

    def test_cell_edit_reaches_the_other_session(self) -> None:
        rows = [_row("001"), _row("002")]
        self.assertTrue(self.alice.start(self.file_path, "アリス", rows))
        self.assertTrue(self.bob.start(self.file_path, "ボブ", rows))
        received = self._received(self.bob)

        edited = [list(row) for row in rows]
        edited[0][COLUMN_MEMO] = "リテイク"
        self.alice.publish_rows(edited)
        self.bob.poll()

        self.assertEqual(len(received), 1)
        merged, cell_updates, structural = apply_diff(rows, received[0])
        self.assertFalse(structural)
        self.assertEqual(cell_updates, [(0, COLUMN_MEMO, "リテイク")])
        self.assertEqual(merged[0][COLUMN_MEMO], "リテイク")

    def test_own_changes_are_not_echoed_back(self) -> None:
        rows = [_row("001")]
        self.alice.start(self.file_path, "アリス", rows)
        received = self._received(self.alice)

        edited = [list(row) for row in rows]
        edited[0][COLUMN_MEMO] = "自分の編集"
        self.alice.publish_rows(edited)
        self.alice.poll()

        self.assertEqual(received, [])

    def test_each_change_is_delivered_once(self) -> None:
        rows = [_row("001")]
        self.alice.start(self.file_path, "アリス", rows)
        self.bob.start(self.file_path, "ボブ", rows)
        received = self._received(self.bob)

        edited = [list(row) for row in rows]
        edited[0][COLUMN_MEMO] = "一度だけ"
        self.alice.publish_rows(edited)
        self.bob.poll()
        self.bob.poll()

        self.assertEqual(len(received), 1)

    def test_changes_made_before_joining_are_not_replayed(self) -> None:
        rows = [_row("001")]
        self.alice.start(self.file_path, "アリス", rows)
        edited = [list(row) for row in rows]
        edited[0][COLUMN_MEMO] = "参加前の変更"
        self.alice.publish_rows(edited)

        self.bob.start(self.file_path, "ボブ", edited)
        received = self._received(self.bob)
        self.bob.poll()

        self.assertEqual(received, [])

    def test_row_insert_and_delete_are_shared(self) -> None:
        rows = [_row("001")]
        self.alice.start(self.file_path, "アリス", rows)
        self.bob.start(self.file_path, "ボブ", rows)
        received = self._received(self.bob)

        self.alice.publish_rows([_row("001"), _row("002")])
        self.bob.poll()
        merged, _cells, structural = apply_diff(rows, received[-1])
        self.assertTrue(structural)
        self.assertEqual([row[COLUMN_CUT_NUMBER] for row in merged], ["001", "002"])

        self.alice.publish_rows([_row("002")])
        self.bob.poll()
        merged, _cells, structural = apply_diff(merged, received[-1])
        self.assertTrue(structural)
        self.assertEqual([row[COLUMN_CUT_NUMBER] for row in merged], ["002"])

    def test_peer_cursor_is_visible_to_the_other_session(self) -> None:
        rows = [_row("001"), _row("002")]
        self.alice.start(self.file_path, "アリス", rows)
        self.bob.start(self.file_path, "ボブ", rows)

        self.alice.publish_cursor("002", "", COLUMN_MEMO)
        self.bob.poll()

        peers = self.bob.peers()
        self.assertEqual([peer.name for peer in peers], ["アリス"])
        self.assertEqual(peers[0].row_key, ("002", ""))
        self.assertEqual(peers[0].column, COLUMN_MEMO)
        self.assertEqual(peers[0].color, self.alice.color())

    def test_leaving_removes_the_peer(self) -> None:
        rows = [_row("001")]
        self.alice.start(self.file_path, "アリス", rows)
        self.bob.start(self.file_path, "ボブ", rows)
        self.bob.poll()
        self.assertEqual(len(self.bob.peers()), 1)

        self.alice.stop()
        self.bob.poll()
        self.assertEqual(self.bob.peers(), [])

    def test_adopted_rows_are_not_published_again(self) -> None:
        rows = [_row("001")]
        self.alice.start(self.file_path, "アリス", rows)
        self.bob.start(self.file_path, "ボブ", rows)
        received_by_alice = self._received(self.alice)

        edited = [list(row) for row in rows]
        edited[0][COLUMN_MEMO] = "ボブの編集"
        self.bob.publish_rows(edited)
        self.alice.poll()
        self.assertEqual(len(received_by_alice), 1)

        # 受信した内容を取り込んだ側は、それを送り返さない。
        self.alice.adopt_rows(edited)
        received_by_bob = self._received(self.bob)
        self.alice.publish_rows(edited)
        self.bob.poll()
        self.assertEqual(received_by_bob, [])


if __name__ == "__main__":
    unittest.main()
