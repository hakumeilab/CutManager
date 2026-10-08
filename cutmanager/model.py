from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
import re
from typing import NamedTuple

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from .constants import (
    BG_STATE_APPROVED,
    BG_STATE_RAW,
    COLUMN_AB_GROUP,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_BG_STATE,
    COLUMN_CUT_NUMBER,
    COLUMN_STATUS,
    COLUMN_THUMBNAIL,
    COLUMN_TP_LOAD_COUNT,
    COLUMN_TP_STATE,
    COLUMN_VIDEO_PATH,
    CSV_HEADERS,
    STATUS_OPTIONS,
    TP_STATE_CHECKED,
    TP_STATE_UNCHECKED,
)

# ユーザーが直接編集できない列（プログラムが管理する）。
READONLY_COLUMNS = frozenset({COLUMN_VIDEO_PATH, COLUMN_THUMBNAIL})
# コピー/貼り付け/クリア/一括入力の対象外にする列。
NON_DATA_COLUMNS = frozenset({COLUMN_THUMBNAIL})
from .folder_import import make_cut_key
from .history import HistoryCommand, HistoryManager


SORT_TOKEN_PATTERN = re.compile(r"\d+|\D+")
STATUS_SHARED = STATUS_OPTIONS[1]
STATUS_BANK = STATUS_OPTIONS[2]
STATUS_MISSING = STATUS_OPTIONS[3]

# data() は描画のたびに 1 セルあたり十数回呼ばれる。PySide6 の短縮名
# （Qt.DisplayRole など）は 1 回の参照に数マイクロ秒かかるため、整数に直して比較する。
_DISPLAY_ROLE = int(Qt.ItemDataRole.DisplayRole)
_EDIT_ROLE = int(Qt.ItemDataRole.EditRole)
_DECORATION_ROLE = int(Qt.ItemDataRole.DecorationRole)
_BACKGROUND_ROLE = int(Qt.ItemDataRole.BackgroundRole)
_FOREGROUND_ROLE = int(Qt.ItemDataRole.ForegroundRole)
_TOOLTIP_ROLE = int(Qt.ItemDataRole.ToolTipRole)
_HANDLED_DATA_ROLES = frozenset(
    {_DISPLAY_ROLE, _EDIT_ROLE, _DECORATION_ROLE, _BACKGROUND_ROLE, _FOREGROUND_ROLE}
)
_TEXT_ROLES = [_DISPLAY_ROLE, _EDIT_ROLE]
_COLOR_ROLES = [_BACKGROUND_ROLE, _FOREGROUND_ROLE]
_TEXT_AND_COLOR_ROLES = _TEXT_ROLES + _COLOR_ROLES
_DECORATION_ROLES = [_DECORATION_ROLE]

_COLUMN_COUNT = len(CSV_HEADERS)
_STR_TYPE_SET = {str}
# これより多くの行が一度に変わったら、行ごとではなく範囲まとめて再描画を通知する。
_BULK_CHANGE_ROW_THRESHOLD = 32

_READONLY_FLAGS = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
_EDITABLE_FLAGS = _READONLY_FLAGS | Qt.ItemFlag.ItemIsEditable
_NO_FLAGS = Qt.ItemFlag.NoItemFlags

# 変わると行全体の配色が変わりうる列。
_ROW_STYLE_COLUMNS = frozenset(
    {COLUMN_STATUS, COLUMN_TP_LOAD_COUNT, COLUMN_BG_LOAD_COUNT, COLUMN_TP_STATE, COLUMN_BG_STATE}
)
# 行の配色とは別に、セル単位で色を上書きする列。
_SPECIAL_COLOR_COLUMNS = (COLUMN_TP_LOAD_COUNT, COLUMN_TP_STATE, COLUMN_BG_LOAD_COUNT, COLUMN_BG_STATE)

_DARK_ROW_COLORS = (QColor("#0f172a"), QColor("#162033"))
_LIGHT_ALTERNATE_ROW_COLOR = QColor("#f7faff")
_DARK_STATUS_ACCENTS = {
    STATUS_SHARED: QColor("#22c55e"),
    STATUS_BANK: QColor("#ef4444"),
    STATUS_MISSING: QColor("#1e3a8a"),
}
_LIGHT_STATUS_ACCENTS = {
    STATUS_SHARED: QColor("#22c55e"),
    STATUS_BANK: QColor("#ef4444"),
    STATUS_MISSING: QColor("#64748b"),
}
_DARK_STATUS_MIX = {
    STATUS_SHARED: 0.22,
    STATUS_BANK: 0.30,
    STATUS_MISSING: 0.50,
}
# TP/BG 状態セルの色。完了側は緑、未完了側はアンバーで塗り分ける。
_MATERIAL_STATE_ACCENTS = {
    TP_STATE_CHECKED: QColor("#22c55e"),
    TP_STATE_UNCHECKED: QColor("#f59e0b"),
    BG_STATE_APPROVED: QColor("#22c55e"),
    BG_STATE_RAW: QColor("#f59e0b"),
}


class _RowStyle(NamedTuple):
    background: QColor | None
    foreground: QColor | None
    # 列 → (背景, 文字色)。行の配色より優先する。
    special: dict[int, tuple[QColor, QColor]]


@dataclass(frozen=True, slots=True)
class CellChange:
    row: int
    column: int
    old_value: str
    new_value: str


class CellChangesCommand(HistoryCommand):
    def __init__(self, model: "CutTableModel", changes: list[CellChange]) -> None:
        self._model = model
        self._changes = list(changes)

    def redo(self) -> None:
        self._model._apply_cell_changes_internal(self._changes, use_new_values=True)

    def undo(self) -> None:
        self._model._apply_cell_changes_internal(self._changes, use_new_values=False)


class RowsSnapshotCommand(HistoryCommand):
    def __init__(
        self,
        model: "CutTableModel",
        old_rows: list[list[str]],
        new_rows: list[list[str]],
        changed_columns: list[int] | None = None,
    ) -> None:
        self._model = model
        # 行そのものは書き換えず必ず差し替えるため、履歴は行を共有して持てる。
        # 全行を複製しないので、行数が多いファイルでも履歴が重くならない。
        self._old_rows = list(old_rows)
        self._new_rows = list(new_rows)
        self._changed_columns = [] if changed_columns is None else list(changed_columns)

    def redo(self) -> None:
        self._model._replace_rows_internal(self._new_rows, self._changed_columns)

    def undo(self) -> None:
        self._model._replace_rows_internal(self._old_rows, self._changed_columns)


class CutTableModel(QAbstractTableModel):
    modifiedChanged = Signal(bool)
    actualRowCountChanged = Signal(int)
    contentChanged = Signal(list)

    def __init__(self, rows: list[list[str]] | None = None, parent=None) -> None:
        super().__init__(parent)
        self._rows = [self._normalize_row(row) for row in (rows or [])]
        self._modified = False
        self._history: HistoryManager | None = None
        self._row_style_cache: dict[int, _RowStyle] = {}
        self._palette_context_cache: tuple[QPalette, bool] | None = None
        self._thumbnail_provider = None
        # 動画パス（casefold）→ 行番号リストの索引。サムネイル更新照合を O(1) にする。
        self._video_path_rows: dict[str, list[int]] | None = None

    def set_history_manager(self, history: HistoryManager | None) -> None:
        self._history = history

    def set_thumbnail_provider(self, provider) -> None:
        self._thumbnail_provider = provider

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._rows) if self._rows else 1

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(CSV_HEADERS)

    def data(self, index: QModelIndex, role: int = _DISPLAY_ROLE):
        # 描画時はフォントや配置など扱わないロールも問い合わせが来るので、先に弾く。
        if role not in _HANDLED_DATA_ROLES or not index.isValid():
            return None

        row = index.row()
        rows = self._rows
        if row >= len(rows):
            return "" if role == _DISPLAY_ROLE or role == _EDIT_ROLE else None

        column = index.column()
        if role == _DISPLAY_ROLE or role == _EDIT_ROLE:
            if column == COLUMN_THUMBNAIL:
                return ""
            return rows[row][column]

        if role == _BACKGROUND_ROLE:
            style = self._row_style(row)
            special = style.special.get(column)
            return style.background if special is None else special[0]

        if role == _FOREGROUND_ROLE:
            style = self._row_style(row)
            special = style.special.get(column)
            return style.foreground if special is None else special[1]

        if role == _DECORATION_ROLE and column == COLUMN_THUMBNAIL:
            return self._thumbnail_for_row(row)

        return None

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
        if role != _EDIT_ROLE or not index.isValid():
            return False

        if index.column() in READONLY_COLUMNS:
            return False

        text = "" if value is None else str(value)
        if self._is_virtual_row(index.row()):
            if text == "":
                return False

            new_rows = list(self._rows)
            appended_row = self._blank_row()
            appended_row[index.column()] = text
            new_rows.append(appended_row)
            self._apply_rows_snapshot(new_rows, modified=True, changed_columns=[index.column()], normalized=True)
            return True

        current_value = self._rows[index.row()][index.column()]
        if current_value == text:
            return False

        return self.apply_cell_changes([(index.row(), index.column(), text)]) > 0

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        if not index.isValid():
            return _NO_FLAGS
        if index.column() in READONLY_COLUMNS:
            return _READONLY_FLAGS
        return _EDITABLE_FLAGS

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = _DISPLAY_ROLE):
        if orientation == Qt.Orientation.Horizontal:
            if role == _DISPLAY_ROLE:
                if 0 <= section < len(CSV_HEADERS):
                    return CSV_HEADERS[section]
                return None
            if role == _TOOLTIP_ROLE:
                return "列見出しをクリックで並べ替え、右端の漏斗ボタンで絞り込みできます。"
            return None

        if role != _DISPLAY_ROLE:
            return None

        return str(section + 1)

    def replace_rows(
        self,
        rows: list[list[str]],
        modified: bool = False,
        *,
        sort_column: int = COLUMN_CUT_NUMBER,
        sort_order: Qt.SortOrder = Qt.SortOrder.AscendingOrder,
    ) -> None:
        normalized_rows = [self._normalize_row(row) for row in rows]
        self._sort_row_list(normalized_rows, sort_column, sort_order)
        self._apply_rows_snapshot(normalized_rows, modified=modified, normalized=True)

    def insert_blank_row(self, position: int | None = None) -> QModelIndex:
        actual_count = len(self._rows)
        insert_at = actual_count if position is None else max(0, min(position, actual_count))
        new_rows = list(self._rows)
        new_rows.insert(insert_at, self._blank_row())
        self._apply_rows_snapshot(new_rows, modified=True, normalized=True)
        return self.index(insert_at, 0)

    def append_rows(self, rows: list[list[str]]) -> None:
        if not rows:
            return
        new_rows = list(self._rows)
        new_rows.extend(self._normalize_row(row) for row in rows)
        self._apply_rows_snapshot(new_rows, modified=True, normalized=True)

    def remove_rows_by_numbers(self, row_numbers: list[int]) -> int:
        targets = sorted({row for row in row_numbers if 0 <= row < len(self._rows)})
        if not targets:
            return 0

        target_set = set(targets)
        new_rows = [row for index, row in enumerate(self._rows) if index not in target_set]
        self._apply_rows_snapshot(new_rows, modified=True, normalized=True)
        return len(targets)

    def clear_indexes(self, indexes: list[QModelIndex]) -> int:
        changes: list[tuple[int, int, str]] = []
        seen: set[tuple[int, int]] = set()

        for index in indexes:
            if not index.isValid() or self._is_virtual_row(index.row()):
                continue
            key = (index.row(), index.column())
            if key in seen:
                continue
            seen.add(key)
            if index.column() in NON_DATA_COLUMNS:
                continue
            if self._rows[index.row()][index.column()] == "":
                continue
            changes.append((index.row(), index.column(), ""))

        return self.apply_cell_changes(changes)

    def apply_cell_changes(self, changes: list[tuple[int, int, str]]) -> int:
        prepared_changes = self._prepare_cell_changes(changes)
        if not prepared_changes:
            return 0

        if self._history is not None:
            self._history.push(CellChangesCommand(self, prepared_changes))
        else:
            self._apply_cell_changes_internal(prepared_changes, use_new_values=True)
            self.set_modified(True)

        return len(prepared_changes)

    def apply_remote_cell_changes(self, changes: list[tuple[int, int, str]]) -> int:
        """共同編集で受信したセル変更を、履歴に積まずに反映する。

        取り消し履歴は各自のローカル操作のためのものなので、他者の変更は
        Undo 対象にしない。素材状態の連動処理も送信側で済んでいるため行わない。
        """

        prepared: list[CellChange] = []
        seen: set[tuple[int, int]] = set()
        for row, column, value in changes:
            if not 0 <= column < len(CSV_HEADERS) or column in NON_DATA_COLUMNS:
                continue
            if not 0 <= row < len(self._rows):
                continue
            key = (row, column)
            if key in seen:
                continue
            seen.add(key)
            new_value = "" if value is None else str(value)
            old_value = self._rows[row][column]
            if old_value == new_value:
                continue
            prepared.append(CellChange(row=row, column=column, old_value=old_value, new_value=new_value))

        if not prepared:
            return 0

        self._apply_cell_changes_internal(prepared, use_new_values=True)
        self.set_modified(True)
        return len(prepared)

    def apply_remote_rows(self, rows: list[list[str]]) -> None:
        """共同編集で受信した行構成を、履歴に積まずに反映する。"""

        self._replace_rows_internal([self._normalize_row(row) for row in rows])
        self.set_modified(True)

    def rows(self) -> list[list[str]]:
        return [row.copy() for row in self._rows]

    def row_at(self, row: int) -> list[str] | None:
        """1 行を読み取り専用で返す。範囲外なら ``None``。"""

        if not 0 <= row < len(self._rows):
            return None
        return self._rows[row]

    def iter_rows(self):
        """内部の行をそのまま順に返す（読み取り専用）。

        集計や同期のように「読むだけ」の処理で、全行コピーを避けるために使う。
        返された行を書き換えてはいけない。
        """

        return iter(self._rows)

    def unique_column_values(self, column: int) -> list[str]:
        if not 0 <= column < len(CSV_HEADERS):
            return []
        return sorted({row[column] for row in self._rows}, key=self._sort_key)

    def cut_keys(self) -> set[tuple[str, str]]:
        return {
            make_cut_key(row[COLUMN_CUT_NUMBER], row[COLUMN_AB_GROUP])
            for row in self._rows
            if row and row[COLUMN_CUT_NUMBER]
        }

    def actual_row_count(self) -> int:
        return len(self._rows)

    def refresh_colors(self) -> None:
        self._clear_color_cache()
        if not self._rows:
            return
        top_left = self.index(0, 0)
        bottom_right = self.index(len(self._rows) - 1, len(CSV_HEADERS) - 1)
        self.dataChanged.emit(
            top_left,
            bottom_right,
            _COLOR_ROLES,
        )

    def is_modified(self) -> bool:
        return self._modified

    def set_modified(self, modified: bool) -> None:
        if self._modified == modified:
            return
        self._modified = modified
        self.modifiedChanged.emit(modified)

    def sort(self, column: int, order: Qt.SortOrder = Qt.SortOrder.AscendingOrder) -> None:
        if column < 0 or column >= len(CSV_HEADERS) or len(self._rows) <= 1:
            return

        self.beginResetModel()
        self._sort_row_list(self._rows, column, order)
        self.endResetModel()

    def _apply_rows_snapshot(
        self,
        new_rows: list[list[str]],
        *,
        modified: bool,
        changed_columns: list[int] | None = None,
        normalized: bool = False,
    ) -> None:
        normalized_rows = new_rows if normalized else [self._normalize_row(row) for row in new_rows]
        if modified and self._history is not None:
            self._history.push(RowsSnapshotCommand(self, self._rows, normalized_rows, changed_columns))
            return

        self._replace_rows_internal(normalized_rows, changed_columns)
        self.set_modified(modified)

    def _replace_rows_internal(self, rows: list[list[str]], changed_columns: list[int] | None = None) -> None:
        self.beginResetModel()
        # 呼び出し元で正規化済み。並べ替えはリストをその場で書き換えるので、
        # 履歴と共有しないようリストだけ作り直す（行そのものは共有する）。
        self._rows = list(rows)
        self._clear_color_cache()
        self._video_path_rows = None
        self.endResetModel()
        self.actualRowCountChanged.emit(len(self._rows))
        if changed_columns:
            self.contentChanged.emit(sorted(set(changed_columns)))

    def _with_material_state_side_effects(
        self, changes: list[tuple[int, int, str]]
    ) -> list[tuple[int, int, str]]:
        """TP/BG 状態の変更に合わせて入れ回数を仮素材 / 本番素材へ合わせる。

        仮素材（未検査 / 素上がり）は入れ回数 0、本番素材（検査済み / 演出OK）は
        入れ回数 1 から数えるため、0 と 1 の間だけ自動で切り替える。
        リテイク（2 以降）や入れ回数を同時に編集した場合は手入力を優先して触らない。
        """

        extended = list(changes)
        explicit_cells = {(row, column) for row, column, _ in changes}
        state_columns = {
            COLUMN_TP_STATE: (COLUMN_TP_LOAD_COUNT, TP_STATE_UNCHECKED, "BGOnly"),
            COLUMN_BG_STATE: (COLUMN_BG_LOAD_COUNT, BG_STATE_RAW, "全セル"),
        }

        for row, column, value in changes:
            target = state_columns.get(column)
            if target is None or not 0 <= row < len(self._rows):
                continue
            count_column, provisional_state, special_value = target
            if (row, count_column) in explicit_cells:
                continue
            current_count = self._rows[row][count_column].strip()
            if current_count == special_value:
                continue
            new_state = str(value or "").strip()
            if new_state == provisional_state:
                if current_count == "1":
                    extended.append((row, count_column, "0"))
            elif new_state and current_count in ("", "0"):
                extended.append((row, count_column, "1"))

        return extended

    def _prepare_cell_changes(self, changes: list[tuple[int, int, str]]) -> list[CellChange]:
        prepared: list[CellChange] = []
        seen: set[tuple[int, int]] = set()

        for row, column, value in self._with_material_state_side_effects(changes):
            if not 0 <= column < len(CSV_HEADERS):
                continue
            if column in NON_DATA_COLUMNS:
                continue
            if not 0 <= row < len(self._rows):
                continue
            key = (row, column)
            if key in seen:
                continue
            seen.add(key)

            new_value = "" if value is None else str(value)
            old_value = self._rows[row][column]
            if old_value == new_value:
                continue
            prepared.append(CellChange(row=row, column=column, old_value=old_value, new_value=new_value))

        return prepared

    def _apply_cell_changes_internal(self, changes: list[CellChange], *, use_new_values: bool) -> None:
        changed_cells: dict[int, set[int]] = defaultdict(set)
        changed_columns: set[int] = set()
        rows_requiring_full_repaint: set[int] = set()

        copied_rows: set[int] = set()
        for change in changes:
            if not 0 <= change.row < len(self._rows):
                continue
            value = change.new_value if use_new_values else change.old_value
            if self._rows[change.row][change.column] == value:
                continue
            if change.row not in copied_rows:
                # 履歴やスナップショットと行を共有しているので、書き換える前に複製する。
                self._rows[change.row] = list(self._rows[change.row])
                copied_rows.add(change.row)
            self._rows[change.row][change.column] = value
            changed_cells[change.row].add(change.column)
            changed_columns.add(change.column)
            if change.column in _ROW_STYLE_COLUMNS:
                rows_requiring_full_repaint.add(change.row)
                self._clear_row_color_cache(change.row)

        if not changed_cells:
            return

        if COLUMN_VIDEO_PATH in changed_columns:
            self._video_path_rows = None

        if len(changed_cells) > _BULK_CHANGE_ROW_THRESHOLD:
            # 行ごとに通知すると、プロキシとビューの処理が行数ぶん繰り返されて重い。
            # 変わった範囲を 1 回で通知し、ビューにはまとめて再描画させる。
            if rows_requiring_full_repaint:
                left_column, right_column, roles = 0, _COLUMN_COUNT - 1, _TEXT_AND_COLOR_ROLES
            else:
                left_column = min(min(columns) for columns in changed_cells.values())
                right_column = max(max(columns) for columns in changed_cells.values())
                roles = _TEXT_ROLES
            self.dataChanged.emit(
                self.index(min(changed_cells), left_column),
                self.index(max(changed_cells), right_column),
                roles,
            )
            self.contentChanged.emit(sorted(changed_columns))
            return

        for row, columns in changed_cells.items():
            if row in rows_requiring_full_repaint:
                left_column = 0
                right_column = len(CSV_HEADERS) - 1
                roles = _TEXT_AND_COLOR_ROLES
            else:
                left_column = min(columns)
                right_column = max(columns)
                roles = _TEXT_ROLES
            self.dataChanged.emit(
                self.index(row, left_column),
                self.index(row, right_column),
                roles,
            )

        self.contentChanged.emit(sorted(changed_columns))

    def _thumbnail_for_row(self, row: int):
        if self._thumbnail_provider is None or not 0 <= row < len(self._rows):
            return None
        video_path = self._rows[row][COLUMN_VIDEO_PATH].strip()
        if not video_path:
            return None
        return self._thumbnail_provider.thumbnail(video_path)

    def video_path_for_row(self, row: int) -> str:
        if not 0 <= row < len(self._rows):
            return ""
        return self._rows[row][COLUMN_VIDEO_PATH].strip()

    def refresh_thumbnails_for_path(self, video_path: str) -> None:
        """指定パスに一致するサムネイルセルの再描画を促す。

        provider が渡すパスも行に保持されたパスも import 時点で解決済みの絶対パスの
        ため resolve() はせず、事前構築した索引で O(1) に該当行を引く。
        """
        target = str(video_path or "").strip().casefold()
        if not target:
            return
        rows = self._video_path_row_map().get(target)
        if not rows:
            return
        for row in rows:
            if 0 <= row < len(self._rows):
                cell = self.index(row, COLUMN_THUMBNAIL)
                self.dataChanged.emit(cell, cell, _DECORATION_ROLES)

    def _video_path_row_map(self) -> dict[str, list[int]]:
        if self._video_path_rows is None:
            mapping: dict[str, list[int]] = {}
            for row, row_values in enumerate(self._rows):
                current = row_values[COLUMN_VIDEO_PATH].strip()
                if current:
                    mapping.setdefault(current.casefold(), []).append(row)
            self._video_path_rows = mapping
        return self._video_path_rows

    def _is_virtual_row(self, row: int) -> bool:
        return row >= len(self._rows)

    @staticmethod
    def _blank_row() -> list[str]:
        return [""] * len(CSV_HEADERS)

    @classmethod
    def _normalize_row(cls, values: list[str]) -> list[str]:
        if cls._is_normalized(values):
            # 既に整った行は複製せずそのまま使う（行は書き換えずに差し替える）。
            return values
        normalized = cls._blank_row()
        for index in range(min(len(values), len(CSV_HEADERS))):
            normalized[index] = "" if values[index] is None else str(values[index])
        return normalized

    @staticmethod
    def _is_normalized(values) -> bool:
        return (
            type(values) is list
            and len(values) == _COLUMN_COUNT
            and set(map(type, values)) == _STR_TYPE_SET
        )

    @classmethod
    def _sort_row_list(cls, rows: list[list[str]], column: int, order: Qt.SortOrder) -> None:
        if len(rows) <= 1:
            return

        reverse = order == Qt.SortOrder.DescendingOrder
        if column == COLUMN_CUT_NUMBER:
            sort_key = cls._default_row_sort_key
        else:
            sort_key = lambda row: (cls._sort_key(row[column]), cls._default_row_sort_key(row))
        rows.sort(key=sort_key, reverse=reverse)

    @staticmethod
    def _sort_key(value: str) -> tuple:
        text = str(value or "").strip()
        if not text:
            return (1, ())

        normalized = text.casefold()
        tokens = []
        for chunk in SORT_TOKEN_PATTERN.findall(normalized):
            if chunk.isdigit():
                tokens.append((0, int(chunk)))
            else:
                tokens.append((1, chunk))
        return (0, tuple(tokens), normalized)

    @classmethod
    def _default_row_sort_key(cls, row: list[str]) -> tuple:
        return (
            cls._sort_key(row[COLUMN_CUT_NUMBER]),
            cls._sort_key(row[COLUMN_AB_GROUP]),
        )

    def _row_style(self, row: int) -> _RowStyle:
        """行の配色（行全体の背景/文字色と、特殊セルの上書き色）をまとめて返す。

        描画のたびに 1 セルあたり何度も呼ばれるため、行単位でまとめて計算して
        キャッシュする。パレット判定もキャッシュ世代ごとに 1 回だけ行う。
        """

        style = self._row_style_cache.get(row)
        if style is None:
            style = self._compute_row_style(row)
            self._row_style_cache[row] = style
        return style

    def _palette_context(self) -> tuple[QPalette, bool]:
        context = self._palette_context_cache
        if context is None:
            palette = QApplication.palette()
            context = (palette, self._is_dark_palette(palette))
            self._palette_context_cache = context
        return context

    def _compute_row_style(self, row: int) -> _RowStyle:
        palette, dark = self._palette_context()
        values = self._rows[row]
        base_color = self._base_row_color_for(row, palette, dark)

        status = values[COLUMN_STATUS].strip()
        background = None
        foreground = None
        accent_color = self._status_accent_color_for(status, dark)
        if accent_color is not None:
            background = self._blend_colors(base_color, accent_color, self._status_mix_ratio_for(status, dark))
            if status == STATUS_MISSING:
                foreground = self._contrast_text_color(background, palette)

        special: dict[int, tuple[QColor, QColor]] = {}
        for column in _SPECIAL_COLOR_COLUMNS:
            cell_background = self._special_cell_background(values, column, base_color, dark)
            if cell_background is not None:
                special[column] = (cell_background, self._contrast_text_color(cell_background, palette))
        return _RowStyle(background, foreground, special)

    def _special_cell_background(
        self,
        values: list[str],
        column: int,
        base_color: QColor,
        dark: bool,
    ) -> QColor | None:
        value = values[column].strip()
        if (column == COLUMN_TP_LOAD_COUNT and value == "BGOnly") or (
            column == COLUMN_BG_LOAD_COUNT and value == "全セル"
        ):
            accent_color = self._status_accent_color_for(STATUS_MISSING, dark)
            return self._blend_colors(base_color, accent_color, self._status_mix_ratio_for(STATUS_MISSING, dark))

        if column not in (COLUMN_TP_STATE, COLUMN_BG_STATE):
            return None
        accent_color = _MATERIAL_STATE_ACCENTS.get(value)
        if accent_color is None:
            return None
        return self._blend_colors(base_color, accent_color, 0.30 if dark else 0.20)

    @classmethod
    def _contrast_text_color(cls, background: QColor, palette: QPalette) -> QColor:
        if cls._is_color_dark(background):
            return palette.color(QPalette.ColorRole.BrightText)
        return palette.color(QPalette.ColorRole.Text)

    def _clear_color_cache(self) -> None:
        self._row_style_cache.clear()
        self._palette_context_cache = None

    def _clear_row_color_cache(self, row: int) -> None:
        self._row_style_cache.pop(row, None)

    @staticmethod
    def _base_row_color_for(row: int, palette: QPalette, dark: bool) -> QColor:
        if dark:
            # Reuse the docs dark palette so desktop and web mock feel consistent.
            return _DARK_ROW_COLORS[row % 2]
        if row % 2 == 0:
            return palette.color(QPalette.ColorRole.Base)
        return _LIGHT_ALTERNATE_ROW_COLOR

    @staticmethod
    def _status_accent_color_for(status: str, dark: bool) -> QColor | None:
        return (_DARK_STATUS_ACCENTS if dark else _LIGHT_STATUS_ACCENTS).get(status)

    @staticmethod
    def _status_mix_ratio_for(status: str, dark: bool) -> float:
        if not dark:
            return 0.28 if status == STATUS_MISSING else 0.18
        return _DARK_STATUS_MIX.get(status, 0.18)

    @staticmethod
    def _blend_colors(base: QColor, overlay: QColor, overlay_alpha: float) -> QColor:
        alpha = max(0.0, min(1.0, overlay_alpha))
        inverse = 1.0 - alpha
        return QColor(
            round((base.red() * inverse) + (overlay.red() * alpha)),
            round((base.green() * inverse) + (overlay.green() * alpha)),
            round((base.blue() * inverse) + (overlay.blue() * alpha)),
        )

    @staticmethod
    def _is_color_dark(color: QColor) -> bool:
        luminance = (0.299 * color.red()) + (0.587 * color.green()) + (0.114 * color.blue())
        return luminance < 128

    @staticmethod
    def _is_dark_palette(palette: QPalette) -> bool:
        app = QApplication.instance()
        if app is not None:
            try:
                if app.styleHints().colorScheme() == Qt.ColorScheme.Dark:
                    return True
                if app.styleHints().colorScheme() == Qt.ColorScheme.Light:
                    return False
            except AttributeError:
                pass
        return CutTableModel._is_color_dark(palette.color(QPalette.ColorRole.Base))
