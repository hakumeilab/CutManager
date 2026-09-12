"""NAS 上の同じプロジェクトファイルを複数人でリアルタイム共有するための同期層。

サーバーを立てずに運用できるよう、プロジェクトファイルの隣に作るサイドカー
フォルダー（``<ファイル名>.cutsync``）だけで同期する。

    cut_list.cutmgr
    cut_list.cutmgr.cutsync/
        peers/<セッションID>.json   … 参加者と選択セル（プレゼンス）
        ops/<セッションID>.json     … そのセッションが発行した変更ログ

各クライアントは一定間隔でサイドカーを読み、自分以外が書いた変更を取り込む。
行の同一性は「カット番号 + AB分け」（``make_cut_key``）で判定する。CutManager が
既に素材取り込みの突き合わせに使っている業務キーで、CSV フォーマットを変えずに
済むためこの方式を採る。カット番号が空の行は同期対象外（キーが作れないため）で、
カット番号が入力された時点で同期に乗る。
"""

from __future__ import annotations

import json
import os
import time
import uuid
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, QTimer, Signal

from .constants import COLUMN_AB_GROUP, COLUMN_CUT_NUMBER, COLUMN_THUMBNAIL, CSV_HEADERS
from .folder_import import make_cut_key


PROTOCOL_VERSION = 1
SYNC_DIR_SUFFIX = ".cutsync"
PEERS_DIR_NAME = "peers"
OPS_DIR_NAME = "ops"

# サイドカーの読み取り間隔。SMB 越しでも負荷が軽く、体感はほぼ即時になる値。
POLL_INTERVAL_MS = 1000
# 自分の在席情報を書き込む間隔。
PRESENCE_INTERVAL_MS = 2000
# この秒数だけ更新が途絶えた参加者は離席扱いにする。
PEER_TIMEOUT_SECONDS = 12.0
# 1 セッションの変更ログに残す最大件数。超えたら古いものから捨てる。
MAX_OPS_PER_SESSION = 400

# サムネイルはローカルのキャッシュパスなので共有しない。
NON_SYNCED_COLUMNS = frozenset({COLUMN_THUMBNAIL})

PEER_COLORS = (
    "#e5484d",
    "#0090ff",
    "#30a46c",
    "#f76b15",
    "#8e4ec6",
    "#c2b800",
    "#00a2c7",
    "#e93d82",
)

RowKey = tuple[str, str]


def row_key(row: list[str]) -> RowKey | None:
    """行の同期キー。カット番号が空の行は同期できないので ``None``。"""

    if not row:
        return None
    cut_number = row[COLUMN_CUT_NUMBER] if len(row) > COLUMN_CUT_NUMBER else ""
    ab_group = row[COLUMN_AB_GROUP] if len(row) > COLUMN_AB_GROUP else ""
    if not str(cut_number).strip():
        return None
    return make_cut_key(cut_number, ab_group)


def snapshot_rows(rows: list[list[str]]) -> dict[RowKey, list[str]]:
    """行リストをキー→行の辞書にする。同キーが複数あるときは先頭を採用する。"""

    snapshot: dict[RowKey, list[str]] = {}
    for row in rows:
        key = row_key(row)
        if key is None or key in snapshot:
            continue
        snapshot[key] = list(row)
    return snapshot


@dataclass(frozen=True, slots=True)
class RowDiff:
    """あるスナップショットから次のスナップショットへの差分。"""

    cells: dict[RowKey, dict[int, str]]
    added: dict[RowKey, list[str]]
    removed: tuple[RowKey, ...]

    def is_empty(self) -> bool:
        return not self.cells and not self.added and not self.removed


def diff_snapshots(
    old: dict[RowKey, list[str]],
    new: dict[RowKey, list[str]],
) -> RowDiff:
    cells: dict[RowKey, dict[int, str]] = {}
    added: dict[RowKey, list[str]] = {}

    for key, new_row in new.items():
        old_row = old.get(key)
        if old_row is None:
            added[key] = list(new_row)
            continue
        changed = {
            column: new_row[column]
            for column in range(len(CSV_HEADERS))
            if column not in NON_SYNCED_COLUMNS
            and _cell(new_row, column) != _cell(old_row, column)
        }
        if changed:
            cells[key] = changed

    removed = tuple(key for key in old if key not in new)
    return RowDiff(cells=cells, added=added, removed=removed)


def apply_diff(rows: list[list[str]], diff: RowDiff) -> tuple[list[list[str]], list[tuple[int, int, str]], bool]:
    """差分を行リストへ取り込む。

    戻り値は ``(新しい行リスト, セル更新リスト, 行構成が変わったか)``。
    セル更新リストは行の増減が無いときに、モデルを作り直さず部分描画するために使う。
    """

    new_rows = [list(row) for row in rows]
    index_by_key: dict[RowKey, int] = {}
    for index, row in enumerate(new_rows):
        key = row_key(row)
        if key is not None and key not in index_by_key:
            index_by_key[key] = index

    cell_updates: list[tuple[int, int, str]] = []
    structural = False

    for key, changes in diff.cells.items():
        index = index_by_key.get(key)
        if index is None:
            continue
        for column, value in sorted(changes.items()):
            if column in NON_SYNCED_COLUMNS or not 0 <= column < len(CSV_HEADERS):
                continue
            if _cell(new_rows[index], column) == value:
                continue
            new_rows[index][column] = value
            cell_updates.append((index, column, value))
            if column in (COLUMN_CUT_NUMBER, COLUMN_AB_GROUP):
                # キーそのものが変わるので、以降の照合をやり直す必要がある。
                structural = True

    removed_keys = set(diff.removed) - set(diff.added)
    if removed_keys:
        kept_rows = [row for row in new_rows if row_key(row) not in removed_keys]
        if len(kept_rows) != len(new_rows):
            new_rows = kept_rows
            structural = True

    for key, row in diff.added.items():
        if key in index_by_key and key not in removed_keys:
            continue
        new_rows.append(list(row))
        structural = True

    if structural:
        cell_updates = []
    return new_rows, cell_updates, structural


def diff_to_payload(diff: RowDiff) -> dict:
    return {
        "cells": [
            [key[0], key[1], [[column, value] for column, value in sorted(changes.items())]]
            for key, changes in diff.cells.items()
        ],
        "added": [[key[0], key[1], list(row)] for key, row in diff.added.items()],
        "removed": [[key[0], key[1]] for key in diff.removed],
    }


def payload_to_diff(payload: dict) -> RowDiff:
    cells: dict[RowKey, dict[int, str]] = {}
    for entry in payload.get("cells") or []:
        try:
            cut_number, ab_group, changes = entry
        except (TypeError, ValueError):
            continue
        parsed = {}
        for change in changes or []:
            try:
                column, value = change
            except (TypeError, ValueError):
                continue
            parsed[int(column)] = "" if value is None else str(value)
        if parsed:
            cells[(str(cut_number), str(ab_group))] = parsed

    added: dict[RowKey, list[str]] = {}
    for entry in payload.get("added") or []:
        try:
            cut_number, ab_group, row = entry
        except (TypeError, ValueError):
            continue
        added[(str(cut_number), str(ab_group))] = ["" if cell is None else str(cell) for cell in row or []]

    removed: list[RowKey] = []
    for entry in payload.get("removed") or []:
        try:
            cut_number, ab_group = entry
        except (TypeError, ValueError):
            continue
        removed.append((str(cut_number), str(ab_group)))

    return RowDiff(cells=cells, added=added, removed=tuple(removed))


@dataclass(frozen=True, slots=True)
class PeerCursor:
    """他の参加者 1 人分の在席情報。"""

    session_id: str
    name: str
    color: str
    cut_number: str
    ab_group: str
    column: int
    updated_at: float

    @property
    def row_key(self) -> RowKey | None:
        if not self.cut_number:
            return None
        return (self.cut_number, self.ab_group)


def peer_color(session_id: str) -> str:
    """セッション ID から安定した表示色を選ぶ。"""

    digest = sum(ord(char) * (index + 1) for index, char in enumerate(session_id))
    return PEER_COLORS[digest % len(PEER_COLORS)]


def sync_dir_for(file_path: str) -> Path:
    target = Path(file_path)
    return target.parent / f"{target.name}{SYNC_DIR_SUFFIX}"


def _cell(row: list[str], column: int) -> str:
    if column < 0 or column >= len(row):
        return ""
    value = row[column]
    return "" if value is None else str(value)


def _write_json_atomic(path: Path, payload: dict) -> None:
    """同じ NAS 上に一時ファイルを書いてから置き換える（途中読みを防ぐ）。"""

    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(f"{path.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        with temp_path.open("w", encoding="utf-8", newline="") as handle:
            json.dump(payload, handle, ensure_ascii=False)
        os.replace(temp_path, path)
    except OSError:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def _read_json(path: Path) -> dict | None:
    try:
        with path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, ValueError):
        return None
    return payload if isinstance(payload, dict) else None


class CollabSession(QObject):
    """サイドカー経由でプロジェクトを共同編集するセッション。"""

    peersChanged = Signal(list)
    remoteDiffReceived = Signal(object)
    statusChanged = Signal(str)
    failed = Signal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.session_id = uuid.uuid4().hex[:12]
        self._display_name = ""
        self._file_path: str | None = None
        self._sync_dir: Path | None = None
        self._active = False
        self._sequence = 0
        self._ops: list[dict] = []
        self._snapshot: dict[RowKey, list[str]] = {}
        self._consumed: dict[str, int] = {}
        self._peers: list[PeerCursor] = []
        self._cursor: tuple[str, str, int] = ("", "", -1)
        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(POLL_INTERVAL_MS)
        self._poll_timer.timeout.connect(self.poll)
        self._presence_timer = QTimer(self)
        self._presence_timer.setInterval(PRESENCE_INTERVAL_MS)
        self._presence_timer.timeout.connect(self._publish_presence)

    # ------------------------------------------------------------------ 状態

    def is_active(self) -> bool:
        return self._active

    def file_path(self) -> str | None:
        return self._file_path

    def peers(self) -> list[PeerCursor]:
        return list(self._peers)

    def display_name(self) -> str:
        return self._display_name

    def color(self) -> str:
        return peer_color(self.session_id)

    # ---------------------------------------------------------------- 開始/終了

    def start(self, file_path: str, display_name: str, rows: list[list[str]]) -> bool:
        self.stop()

        sync_dir = sync_dir_for(file_path)
        try:
            (sync_dir / PEERS_DIR_NAME).mkdir(parents=True, exist_ok=True)
            (sync_dir / OPS_DIR_NAME).mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            self.failed.emit(f"共有フォルダーを作成できませんでした: {exc}")
            return False

        self._file_path = file_path
        self._sync_dir = sync_dir
        self._display_name = display_name or "名無し"
        self._snapshot = snapshot_rows(rows)
        self._sequence = 0
        self._ops = []
        self._consumed = {}
        self._peers = []
        self._active = True

        # 既に置かれている変更ログは「参加より前の履歴」なので、読み飛ばし位置だけ合わせる。
        for session_id, ops in self._read_peer_ops().items():
            if ops:
                self._consumed[session_id] = max(int(op.get("seq", 0)) for op in ops)

        self._publish_presence()
        self._poll_timer.start()
        self._presence_timer.start()
        self.statusChanged.emit("共同編集を開始しました。")
        self.poll()
        return True

    def stop(self) -> None:
        if not self._active:
            return
        self._poll_timer.stop()
        self._presence_timer.stop()
        self._remove_own_presence()
        self._active = False
        self._file_path = None
        self._sync_dir = None
        self._peers = []
        self.peersChanged.emit([])
        self.statusChanged.emit("共同編集を終了しました。")

    # ------------------------------------------------------------------ 発信

    def publish_rows(self, rows: list[list[str]]) -> None:
        """ローカルの変更を差分として配信する。"""

        if not self._active or self._sync_dir is None:
            return

        new_snapshot = snapshot_rows(rows)
        diff = diff_snapshots(self._snapshot, new_snapshot)
        self._snapshot = new_snapshot
        if diff.is_empty():
            return

        self._sequence += 1
        self._ops.append(
            {
                "seq": self._sequence,
                "ts": time.time(),
                "diff": diff_to_payload(diff),
            }
        )
        if len(self._ops) > MAX_OPS_PER_SESSION:
            del self._ops[: len(self._ops) - MAX_OPS_PER_SESSION]

        try:
            _write_json_atomic(
                self._sync_dir / OPS_DIR_NAME / f"{self.session_id}.json",
                {
                    "version": PROTOCOL_VERSION,
                    "session": self.session_id,
                    "name": self._display_name,
                    "ops": self._ops,
                },
            )
        except OSError as exc:
            self.failed.emit(f"変更を共有できませんでした: {exc}")

    def adopt_rows(self, rows: list[list[str]]) -> None:
        """受信した変更を取り込んだ直後など、配信せずに基準だけ更新する。"""

        if self._active:
            self._snapshot = snapshot_rows(rows)

    def publish_cursor(self, cut_number: str, ab_group: str, column: int) -> None:
        cursor = (str(cut_number or ""), str(ab_group or "").upper(), int(column))
        if cursor == self._cursor:
            return
        self._cursor = cursor
        if self._active:
            self._publish_presence()

    # ------------------------------------------------------------------ 受信

    def poll(self) -> None:
        if not self._active or self._sync_dir is None:
            return

        self._refresh_peers()

        merged_cells: dict[RowKey, dict[int, str]] = {}
        merged_added: dict[RowKey, list[str]] = {}
        merged_removed: list[RowKey] = []

        for session_id, ops in self._read_peer_ops().items():
            last_seen = self._consumed.get(session_id, 0)
            highest = last_seen
            for op in sorted(ops, key=lambda item: int(item.get("seq", 0))):
                sequence = int(op.get("seq", 0))
                if sequence <= last_seen:
                    continue
                highest = max(highest, sequence)
                diff = payload_to_diff(op.get("diff") or {})
                for key, changes in diff.cells.items():
                    merged_cells.setdefault(key, {}).update(changes)
                for key, row in diff.added.items():
                    merged_added[key] = row
                    if key in merged_removed:
                        merged_removed.remove(key)
                for key in diff.removed:
                    merged_added.pop(key, None)
                    if key not in merged_removed:
                        merged_removed.append(key)
            if highest > last_seen:
                self._consumed[session_id] = highest

        diff = RowDiff(cells=merged_cells, added=merged_added, removed=tuple(merged_removed))
        if not diff.is_empty():
            self.remoteDiffReceived.emit(diff)

    # ---------------------------------------------------------------- 内部処理

    def _read_peer_ops(self) -> dict[str, list[dict]]:
        if self._sync_dir is None:
            return {}
        ops_dir = self._sync_dir / OPS_DIR_NAME
        result: dict[str, list[dict]] = {}
        try:
            entries = sorted(ops_dir.glob("*.json"))
        except OSError:
            return {}
        for entry in entries:
            session_id = entry.stem
            if session_id == self.session_id:
                continue
            payload = _read_json(entry)
            if not payload or int(payload.get("version", 0)) != PROTOCOL_VERSION:
                continue
            ops = payload.get("ops")
            if isinstance(ops, list):
                result[session_id] = [op for op in ops if isinstance(op, dict)]
        return result

    def _refresh_peers(self) -> None:
        if self._sync_dir is None:
            return
        peers_dir = self._sync_dir / PEERS_DIR_NAME
        now = time.time()
        peers: list[PeerCursor] = []
        try:
            entries = sorted(peers_dir.glob("*.json"))
        except OSError:
            entries = []

        for entry in entries:
            session_id = entry.stem
            if session_id == self.session_id:
                continue
            payload = _read_json(entry)
            if not payload:
                continue
            updated_at = float(payload.get("ts") or 0.0)
            if now - updated_at > PEER_TIMEOUT_SECONDS:
                self._discard_stale_peer(entry, session_id, now - updated_at)
                continue
            peers.append(
                PeerCursor(
                    session_id=session_id,
                    name=str(payload.get("name") or "名無し"),
                    color=str(payload.get("color") or peer_color(session_id)),
                    cut_number=str(payload.get("cut") or ""),
                    ab_group=str(payload.get("ab") or ""),
                    column=int(payload.get("column", -1)),
                    updated_at=updated_at,
                )
            )

        peers.sort(key=lambda peer: (peer.name, peer.session_id))
        if peers != self._peers:
            self._peers = peers
            self.peersChanged.emit(list(peers))

    def _discard_stale_peer(self, entry: Path, session_id: str, idle_seconds: float) -> None:
        # 十分に古い置き土産だけ掃除する（時計ずれで生きている参加者を消さない）。
        if idle_seconds < PEER_TIMEOUT_SECONDS * 10:
            return
        for target in (entry, entry.parent.parent / OPS_DIR_NAME / f"{session_id}.json"):
            try:
                target.unlink(missing_ok=True)
            except OSError:
                pass

    def _publish_presence(self) -> None:
        if not self._active or self._sync_dir is None:
            return
        cut_number, ab_group, column = self._cursor
        try:
            _write_json_atomic(
                self._sync_dir / PEERS_DIR_NAME / f"{self.session_id}.json",
                {
                    "version": PROTOCOL_VERSION,
                    "session": self.session_id,
                    "name": self._display_name,
                    "color": self.color(),
                    "cut": cut_number,
                    "ab": ab_group,
                    "column": column,
                    "ts": time.time(),
                },
            )
        except OSError as exc:
            self.failed.emit(f"在席情報を共有できませんでした: {exc}")

    def _remove_own_presence(self) -> None:
        if self._sync_dir is None:
            return
        try:
            (self._sync_dir / PEERS_DIR_NAME / f"{self.session_id}.json").unlink(missing_ok=True)
        except OSError:
            pass
