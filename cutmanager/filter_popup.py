from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QPushButton, QVBoxLayout


class ColumnFilterPopup(QDialog):
    def __init__(
        self,
        column_name: str,
        values: list[str],
        checked_values: set[str],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._all_values = list(values)

        self.setWindowFlags(Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint)
        self.setWindowTitle(column_name)
        self.setMinimumWidth(260)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        title_label = QLabel(f"{column_name} の絞り込み", self)
        layout.addWidget(title_label)

        filter_tools_layout = QHBoxLayout()
        filter_tools_layout.setSpacing(6)
        select_all_button = QPushButton("全選択", self)
        clear_all_button = QPushButton("全解除", self)
        filter_tools_layout.addWidget(select_all_button)
        filter_tools_layout.addWidget(clear_all_button)
        layout.addLayout(filter_tools_layout)

        self.value_list = QListWidget(self)
        self.value_list.setAlternatingRowColors(True)
        self.value_list.setSelectionMode(QListWidget.SelectionMode.NoSelection)

        for value in values:
            item = QListWidgetItem(self._label_for_value(value))
            item.setData(Qt.ItemDataRole.UserRole, value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            item.setCheckState(Qt.CheckState.Checked if value in checked_values else Qt.CheckState.Unchecked)
            self.value_list.addItem(item)

        layout.addWidget(self.value_list, 1)

        button_box = QDialogButtonBox(self)
        self.apply_button = button_box.addButton("適用", QDialogButtonBox.ButtonRole.AcceptRole)
        self.cancel_button = button_box.addButton("閉じる", QDialogButtonBox.ButtonRole.RejectRole)
        layout.addWidget(button_box)

        select_all_button.clicked.connect(lambda: self._set_all_checked(Qt.CheckState.Checked))
        clear_all_button.clicked.connect(lambda: self._set_all_checked(Qt.CheckState.Unchecked))
        self.apply_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

    def selected_values(self) -> set[str]:
        selected: set[str] = set()
        for index in range(self.value_list.count()):
            item = self.value_list.item(index)
            if item.checkState() == Qt.CheckState.Checked:
                selected.add(str(item.data(Qt.ItemDataRole.UserRole) or ""))
        return selected

    def all_values(self) -> set[str]:
        return set(self._all_values)

    def _set_all_checked(self, state: Qt.CheckState) -> None:
        for index in range(self.value_list.count()):
            self.value_list.item(index).setCheckState(state)

    @staticmethod
    def _label_for_value(value: str) -> str:
        return value if value else "(空欄)"
