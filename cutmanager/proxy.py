from __future__ import annotations

from PySide6.QtCore import QSortFilterProxyModel, Qt


class CutFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._allowed_values_by_column: dict[int, set[str]] = {}
        self.setDynamicSortFilter(True)
        self.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseSensitive)

    def set_allowed_values(self, column: int, values: set[str] | None) -> None:
        if values is None:
            self._allowed_values_by_column.pop(column, None)
        else:
            self._allowed_values_by_column[column] = {str(value) for value in values}
        self.invalidateFilter()

    def clear_allowed_values(self, column: int) -> None:
        if column in self._allowed_values_by_column:
            del self._allowed_values_by_column[column]
            self.invalidateFilter()

    def clear_all_filters(self) -> None:
        if not self._allowed_values_by_column:
            return
        self._allowed_values_by_column.clear()
        self.invalidateFilter()

    def allowed_values(self, column: int) -> set[str] | None:
        values = self._allowed_values_by_column.get(column)
        if values is None:
            return None
        return set(values)

    def has_active_filters(self) -> bool:
        return bool(self._allowed_values_by_column)

    def filtered_columns(self) -> set[int]:
        return set(self._allowed_values_by_column)

    def filterAcceptsRow(self, source_row: int, source_parent) -> bool:
        model = self.sourceModel()
        if model is None:
            return True

        for column, allowed_values in self._allowed_values_by_column.items():
            index = model.index(source_row, column, source_parent)
            value = str(model.data(index, Qt.ItemDataRole.DisplayRole) or "")
            if value not in allowed_values:
                return False

        return True
