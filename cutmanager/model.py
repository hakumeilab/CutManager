from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
import re

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal

from .constants import COLUMN_AB_GROUP, COLUMN_CUT_NUMBER, CSV_HEADERS
from .folder_import import make_cut_key
from .history import HistoryCommand, HistoryManager


SORT_TOKEN_PATTERN = re.compile(r"\d+|\D+")


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

    def set_history_manager(self, history: HistoryManager | None) -> None:
        self._history = history

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

        if role not in (Qt.DisplayRole, Qt.EditRole):
            return None

        if self._is_virtual_row(index.row()):
            return ""

        return self._rows[index.row()][index.column()]

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
        if role != Qt.EditRole or not index.isValid():
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
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

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

        for change in changes:
            if not 0 <= change.row < len(self._rows):
                continue
            value = change.new_value if use_new_values else change.old_value
            if self._rows[change.row][change.column] == value:
                continue
            self._rows[change.row][change.column] = value
            changed_cells[change.row].add(change.column)
            changed_columns.add(change.column)

        if not changed_cells:
            return

        for row, columns in changed_cells.items():
            left_column = min(columns)
            right_column = max(columns)
            self.dataChanged.emit(
                self.index(row, left_column),
                self.index(row, right_column),
                [Qt.DisplayRole, Qt.EditRole],
            )

        self.contentChanged.emit(sorted(changed_columns))

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
