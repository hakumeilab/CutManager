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


PROTOCOL_VERSION = 2
SYNC_DIR_SUFFIX = ".cutsync"
PEERS_DIR_NAME = "peers"
OPS_DIR_NAME = "ops"
# 変更ログは 1 行 1 操作の JSON Lines。追記した分だけを読むため、
# 編集が積み重なっても書き込み量も読み取り量も増えない。
OPS_FILE_EXTENSION = ".jsonl"

# サイドカーの読み取り間隔。SMB 越しでも負荷が軽く、体感はほぼ即時になる値。
POLL_INTERVAL_MS = 1000
# 自分の在席情報を書き込む間隔。
PRESENCE_INTERVAL_MS = 2000
# この秒数だけ更新が途絶えた参加者は離席扱いにする。
PEER_TIMEOUT_SECONDS = 12.0
# 変更ログがこのサイズを超えたら、直近の操作だけ残して書き直す。
OPS_FILE_MAX_BYTES = 256 * 1024
# 書き直すときに残す操作の数。全員が 1 秒間隔で読むので、直近だけあれば足りる。
OPS_ROTATE_KEEP = 50

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

    size = len(row)
    if size <= COLUMN_CUT_NUMBER:
        return None
    cut_number = row[COLUMN_CUT_NUMBER]
    if not cut_number:
        # 空のカット番号が大半なので、strip する前にここで弾く。
        return None
    ab_group = row[COLUMN_AB_GROUP] if size > COLUMN_AB_GROUP else ""
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
        if old_row == new_row:
            # 大半の行は変わらない。列ごとの比較へ入る前にまとめて弾く。
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


def apply_diff(rows, diff: RowDiff) -> tuple[list[list[str]], list[tuple[int, int, str]], bool]:
    """差分を行リストへ取り込む。

    戻り値は ``(新しい行リスト, セル更新リスト, 行構成が変わったか)``。
    セル更新リストは行の増減が無いときに、モデルを作り直さず部分描画するために使う。

    変更しない行は元のリスト要素をそのまま使い回す（書き換える行だけ複製する）。
    全行を複製しないぶん速い代わりに、戻り値の行を直接書き換えてはいけない。
    """

    new_rows = list(rows)
    index_by_key: dict[RowKey, int] = {}
    for index, row in enumerate(new_rows):
        key = row_key(row)
        if key is not None and key not in index_by_key:
            index_by_key[key] = index

    cell_updates: list[tuple[int, int, str]] = []
    structural = False
    copied: set[int] = set()

    for key, changes in diff.cells.items():
        index = index_by_key.get(key)
        if index is None:
            continue
        for column, value in sorted(changes.items()):
            if column in NON_SYNCED_COLUMNS or not 0 <= column < len(CSV_HEADERS):
                continue
            if _cell(new_rows[index], column) == value:
                continue
            if index not in copied:
                new_rows[index] = list(new_rows[index])
                copied.add(index)
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


def _file_size(path: Path) -> int:
    try:
        return path.stat().st_size
    except OSError:
        return 0


def _read_lines_from(path: Path, offset: int) -> tuple[list[str], int]:
    """``offset`` から後ろを読み、行として完成している分だけ返す。

    戻り値は ``(行のリスト, 読み進めたバイト数)``。書き込み途中で末尾が
    途切れていた場合、その分は読み進めず次回に持ち越す。
    """

    try:
        with path.open("rb") as handle:
            handle.seek(offset)
            chunk = handle.read()
    except OSError:
        return [], 0

    if not chunk:
        return [], 0

    end = chunk.rfind(b"\n")
    if end < 0:
        return [], 0

    complete = chunk[: end + 1]
    try:
        text = complete.decode("utf-8")
    except UnicodeDecodeError:
        return [], len(complete)
    return [line for line in text.splitlines() if line.strip()], len(complete)


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
        # 参加者ごとの「どこまで読んだか」（変更ログ内のバイト位置）。
        self._offsets: dict[str, int] = {}
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
        self._offsets = {}
        self._peers = []
        self._active = True

        # 既に置かれている変更ログは「参加より前の履歴」なので、末尾から読み始める。
        for entry in self._ops_entries():
            if entry.stem != self.session_id:
                self._offsets[entry.stem] = _file_size(entry)

        # 自分のログは毎回まっさらから始める（seq もここで 1 に戻る）。
        try:
            self._ops_path().write_text("", encoding="utf-8")
        except OSError as exc:
            self.failed.emit(f"変更ログを初期化できませんでした: {exc}")

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

    def publish_rows(self, rows) -> None:
        """ローカルの変更を差分として配信する。"""

        if not self._active or self._sync_dir is None:
            return

        new_snapshot = snapshot_rows(rows)
        diff = diff_snapshots(self._snapshot, new_snapshot)
        self._snapshot = new_snapshot
        if diff.is_empty():
            return

        self._sequence += 1
        operation = {
            "version": PROTOCOL_VERSION,
            "session": self.session_id,
            "name": self._display_name,
            "seq": self._sequence,
            "ts": time.time(),
            "diff": diff_to_payload(diff),
        }
        self._ops.append(operation)
        if len(self._ops) > OPS_ROTATE_KEEP:
            del self._ops[: len(self._ops) - OPS_ROTATE_KEEP]

        ops_path = self._ops_path()
        try:
            with ops_path.open("a", encoding="utf-8", newline="\n") as handle:
                handle.write(json.dumps(operation, ensure_ascii=False) + "\n")
            self._rotate_ops_if_needed(ops_path)
        except OSError as exc:
            self.failed.emit(f"変更を共有できませんでした: {exc}")

    def _rotate_ops_if_needed(self, ops_path: Path) -> None:
        """ログが膨らんだら直近の操作だけ残して書き直す。"""

        if _file_size(ops_path) <= OPS_FILE_MAX_BYTES:
            return
        recent = self._ops[-OPS_ROTATE_KEEP:]
        body = "".join(json.dumps(operation, ensure_ascii=False) + "\n" for operation in recent)
        temp_path = ops_path.with_name(f"{ops_path.name}.{uuid.uuid4().hex[:8]}.tmp")
        try:
            temp_path.write_text(body, encoding="utf-8", newline="\n")
            os.replace(temp_path, ops_path)
        except OSError:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass

    def adopt_diff(self, diff: RowDiff) -> None:
        """受信して取り込んだ差分を、配信の基準へ反映する。

        全行から作り直すより軽く、次の配信で送り返すことも防げる。
        """

        if not self._active:
            return
        for key, changes in diff.cells.items():
            row = self._snapshot.get(key)
            if row is None:
                continue
            for column, value in changes.items():
                if 0 <= column < len(row):
                    row[column] = value
        for key in diff.removed:
            if key not in diff.added:
                self._snapshot.pop(key, None)
        for key, row in diff.added.items():
            self._snapshot[key] = list(row)

    def adopt_rows(self, rows) -> None:
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

        for session_id, ops in self._read_new_peer_ops().items():
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

    def _ops_path(self, session_id: str | None = None) -> Path:
        assert self._sync_dir is not None
        name = session_id or self.session_id
        return self._sync_dir / OPS_DIR_NAME / f"{name}{OPS_FILE_EXTENSION}"

    def _ops_entries(self) -> list[Path]:
        if self._sync_dir is None:
            return []
        try:
            return sorted((self._sync_dir / OPS_DIR_NAME).glob(f"*{OPS_FILE_EXTENSION}"))
        except OSError:
            return []

    def _read_new_peer_ops(self) -> dict[str, list[dict]]:
        """各参加者の変更ログのうち、前回読んだ位置より後ろだけを読む。"""

        result: dict[str, list[dict]] = {}
        for entry in self._ops_entries():
            session_id = entry.stem
            if session_id == self.session_id:
                continue

            offset = self._offsets.get(session_id, 0)
            size = _file_size(entry)
            if size == offset:
                # 追記が無いので読み込み自体を省く。待機中はここで終わる。
                continue
            if size < offset:
                # ログが書き直された。先頭から読み直し、seq で二重取り込みを防ぐ。
                offset = 0

            lines, consumed_bytes = _read_lines_from(entry, offset)
            self._offsets[session_id] = offset + consumed_bytes
            if not lines:
                continue

            ops: list[dict] = []
            for line in lines:
                try:
                    operation = json.loads(line)
                except ValueError:
                    continue
                if not isinstance(operation, dict):
                    continue
                if int(operation.get("version", 0)) != PROTOCOL_VERSION:
                    continue
                ops.append(operation)
            if ops:
                result[session_id] = ops
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
        for target in (entry, entry.parent.parent / OPS_DIR_NAME / f"{session_id}{OPS_FILE_EXTENSION}"):
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
