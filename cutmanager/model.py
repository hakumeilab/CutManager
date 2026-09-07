from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
import re

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from .constants import (
    COLUMN_AB_GROUP,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_CUT_NUMBER,
    COLUMN_STATUS,
    COLUMN_THUMBNAIL,
    COLUMN_TP_LOAD_COUNT,
    COLUMN_VIDEO_PATH,
    CSV_HEADERS,
    STATUS_OPTIONS,
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
        self._old_rows = [row.copy() for row in old_rows]
        self._new_rows = [row.copy() for row in new_rows]
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
        self._row_background_cache: dict[int, QColor | None] = {}
        self._row_foreground_cache: dict[int, QColor | None] = {}
        self._special_background_cache: dict[tuple[int, int], QColor | None] = {}
        self._special_foreground_cache: dict[tuple[int, int], QColor | None] = {}
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

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None

        if self._is_virtual_row(index.row()):
            return "" if role in (Qt.DisplayRole, Qt.EditRole) else None

        if index.column() == COLUMN_THUMBNAIL:
            if role == Qt.DecorationRole:
                return self._thumbnail_for_row(index.row())
            if role in (Qt.DisplayRole, Qt.EditRole):
                return ""

        if role in (Qt.DisplayRole, Qt.EditRole):
            return self._rows[index.row()][index.column()]

        if role == Qt.BackgroundRole:
            return self._cell_background_color(index.row(), index.column())

        if role == Qt.ForegroundRole:
            return self._cell_foreground_color(index.row(), index.column())

        return None

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid():
            return False

        if index.column() in READONLY_COLUMNS:
            return False

        text = "" if value is None else str(value)
        if self._is_virtual_row(index.row()):
            if text == "":
                return False

            new_rows = self.rows()
            new_rows.append(self._blank_row())
            new_rows[index.row()][index.column()] = text
            self._apply_rows_snapshot(new_rows, modified=True, changed_columns=[index.column()])
            return True

        current_value = self._rows[index.row()][index.column()]
        if current_value == text:
            return False

        return self.apply_cell_changes([(index.row(), index.column(), text)]) > 0

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        if not index.isValid():
            return Qt.NoItemFlags
        base_flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.column() in READONLY_COLUMNS:
            return base_flags
        return base_flags | Qt.ItemIsEditable

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                if 0 <= section < len(CSV_HEADERS):
                    return CSV_HEADERS[section]
                return None
            if role == Qt.ToolTipRole:
                return "列見出しをクリックで並べ替え、右端の漏斗ボタンで絞り込みできます。"
            return None

        if role != Qt.DisplayRole:
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
        self._apply_rows_snapshot(normalized_rows, modified=modified)

    def insert_blank_row(self, position: int | None = None) -> QModelIndex:
        actual_count = len(self._rows)
        insert_at = actual_count if position is None else max(0, min(position, actual_count))
        new_rows = self.rows()
        new_rows.insert(insert_at, self._blank_row())
        self._apply_rows_snapshot(new_rows, modified=True)
        return self.index(insert_at, 0)

    def append_rows(self, rows: list[list[str]]) -> None:
        if not rows:
            return
        new_rows = self.rows()
        new_rows.extend(self._normalize_row(row) for row in rows)
        self._apply_rows_snapshot(new_rows, modified=True)

    def remove_rows_by_numbers(self, row_numbers: list[int]) -> int:
        targets = sorted({row for row in row_numbers if 0 <= row < len(self._rows)})
        if not targets:
            return 0

        target_set = set(targets)
        new_rows = [row.copy() for index, row in enumerate(self._rows) if index not in target_set]
        self._apply_rows_snapshot(new_rows, modified=True)
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

    def rows(self) -> list[list[str]]:
        return [row.copy() for row in self._rows]

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
            [Qt.BackgroundRole, Qt.ForegroundRole],
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
    ) -> None:
        normalized_rows = [self._normalize_row(row) for row in new_rows]
        if modified and self._history is not None:
            self._history.push(RowsSnapshotCommand(self, self.rows(), normalized_rows, changed_columns))
            return

        self._replace_rows_internal(normalized_rows, changed_columns)
        self.set_modified(modified)

    def _replace_rows_internal(self, rows: list[list[str]], changed_columns: list[int] | None = None) -> None:
        self.beginResetModel()
        self._rows = [self._normalize_row(row) for row in rows]
        self._clear_color_cache()
        self._video_path_rows = None
        self.endResetModel()
        self.actualRowCountChanged.emit(len(self._rows))
        if changed_columns:
            self.contentChanged.emit(sorted(set(changed_columns)))

    def _prepare_cell_changes(self, changes: list[tuple[int, int, str]]) -> list[CellChange]:
        prepared: list[CellChange] = []
        seen: set[tuple[int, int]] = set()

        for row, column, value in changes:
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

        for change in changes:
            if not 0 <= change.row < len(self._rows):
                continue
            value = change.new_value if use_new_values else change.old_value
            if self._rows[change.row][change.column] == value:
                continue
            self._rows[change.row][change.column] = value
            changed_cells[change.row].add(change.column)
            changed_columns.add(change.column)
            if change.column in (COLUMN_STATUS, COLUMN_TP_LOAD_COUNT, COLUMN_BG_LOAD_COUNT):
                rows_requiring_full_repaint.add(change.row)
                self._clear_row_color_cache(change.row)

        if not changed_cells:
            return

        if COLUMN_VIDEO_PATH in changed_columns:
            self._video_path_rows = None

        for row, columns in changed_cells.items():
            if row in rows_requiring_full_repaint:
                left_column = 0
                right_column = len(CSV_HEADERS) - 1
                roles = [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole, Qt.ForegroundRole]
            else:
                left_column = min(columns)
                right_column = max(columns)
                roles = [Qt.DisplayRole, Qt.EditRole]
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
                self.dataChanged.emit(cell, cell, [Qt.DecorationRole])

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
        normalized = cls._blank_row()
        for index in range(min(len(values), len(CSV_HEADERS))):
            normalized[index] = "" if values[index] is None else str(values[index])
        return normalized

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

    def _cell_background_color(self, row: int, column: int) -> QColor | None:
        special_background = self._cached_special_count_cell_background(row, column)
        if special_background is not None:
            return special_background
        if row not in self._row_background_cache:
            self._row_background_cache[row] = self._row_background_color(row)
        return self._row_background_cache[row]

    def _cell_foreground_color(self, row: int, column: int) -> QColor | None:
        special_foreground = self._cached_special_count_cell_foreground(row, column)
        if special_foreground is not None:
            return special_foreground
        if row not in self._row_foreground_cache:
            self._row_foreground_cache[row] = self._row_foreground_color(row)
        return self._row_foreground_cache[row]

    def _clear_color_cache(self) -> None:
        self._row_background_cache.clear()
        self._row_foreground_cache.clear()
        self._special_background_cache.clear()
        self._special_foreground_cache.clear()

    def _clear_row_color_cache(self, row: int) -> None:
        self._row_background_cache.pop(row, None)
        self._row_foreground_cache.pop(row, None)
        self._special_background_cache = {key: value for key, value in self._special_background_cache.items() if key[0] != row}
        self._special_foreground_cache = {key: value for key, value in self._special_foreground_cache.items() if key[0] != row}

    def _cached_special_count_cell_background(self, row: int, column: int) -> QColor | None:
        key = (row, column)
        if key not in self._special_background_cache:
            self._special_background_cache[key] = self._special_count_cell_background(row, column)
        return self._special_background_cache[key]

    def _cached_special_count_cell_foreground(self, row: int, column: int) -> QColor | None:
        key = (row, column)
        if key not in self._special_foreground_cache:
            self._special_foreground_cache[key] = self._special_count_cell_foreground(row, column)
        return self._special_foreground_cache[key]

    def _row_background_color(self, row: int) -> QColor | None:
        palette = QApplication.palette()
        base_color = self._base_row_color(row, palette)
        status = self._rows[row][COLUMN_STATUS].strip()
        accent_color = self._status_accent_color(status)
        if accent_color is None:
            return None
        mix_ratio = self._status_mix_ratio(status, palette)
        return self._blend_colors(base_color, accent_color, mix_ratio)

    def _special_count_cell_background(self, row: int, column: int) -> QColor | None:
        if not self._is_special_count_cell(row, column):
            return None
        palette = QApplication.palette()
        base_color = self._base_row_color(row, palette)
        accent_color = self._status_accent_color(STATUS_MISSING)
        if accent_color is None:
            return None
        mix_ratio = self._status_mix_ratio(STATUS_MISSING, palette)
        return self._blend_colors(base_color, accent_color, mix_ratio)

    def _row_foreground_color(self, row: int) -> QColor | None:
        palette = QApplication.palette()
        background = self._row_background_color(row)
        status = self._rows[row][COLUMN_STATUS].strip()
        if status != STATUS_MISSING or background is None:
            return None
        if self._is_color_dark(background):
            return palette.color(QPalette.ColorRole.BrightText)
        return palette.color(QPalette.ColorRole.Text)

    def _special_count_cell_foreground(self, row: int, column: int) -> QColor | None:
        background = self._special_count_cell_background(row, column)
        if background is None:
            return None
        palette = QApplication.palette()
        if self._is_color_dark(background):
            return palette.color(QPalette.ColorRole.BrightText)
        return palette.color(QPalette.ColorRole.Text)

    def _is_special_count_cell(self, row: int, column: int) -> bool:
        if not 0 <= row < len(self._rows):
            return False
        if column == COLUMN_TP_LOAD_COUNT:
            return self._rows[row][column].strip() == "BGOnly"
        if column == COLUMN_BG_LOAD_COUNT:
            return self._rows[row][column].strip() == "全セル"
        return False

    @staticmethod
    def _base_row_color(row: int, palette: QPalette) -> QColor:
        if CutTableModel._is_dark_palette(palette):
            # Reuse the docs dark palette so desktop and web mock feel consistent.
            return QColor("#0f172a" if row % 2 == 0 else "#162033")
        if row % 2 == 0:
            return palette.color(QPalette.ColorRole.Base)
        return QColor("#f7faff")

    @staticmethod
    def _status_accent_color(status: str) -> QColor | None:
        dark_mode = CutTableModel._is_dark_palette(QApplication.palette())
        accent_by_status = (
            {
                STATUS_SHARED: QColor("#22c55e"),
                STATUS_BANK: QColor("#ef4444"),
                STATUS_MISSING: QColor("#1e3a8a"),
            }
            if dark_mode
            else {
                STATUS_SHARED: QColor("#22c55e"),
                STATUS_BANK: QColor("#ef4444"),
                STATUS_MISSING: QColor("#64748b"),
            }
        )
        return accent_by_status.get(status)

    @staticmethod
    def _status_mix_ratio(status: str, palette: QPalette) -> float:
        if not CutTableModel._is_dark_palette(palette):
            return 0.28 if status == STATUS_MISSING else 0.18
        dark_mix = {
            STATUS_SHARED: 0.22,
            STATUS_BANK: 0.30,
            STATUS_MISSING: 0.50,
        }
        return dark_mix.get(status, 0.18)

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
