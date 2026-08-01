from __future__ import annotations

from PySide6.QtGui import QKeySequence
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QKeySequenceEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from .shortcuts import SHORTCUT_SPECS


class SettingsDialog(QDialog):
    def __init__(
        self,
        undo_limit: int,
        shortcuts: dict[str, list[str]] | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("環境設定")

        self.undo_limit_spin = QSpinBox(self)
        self.undo_limit_spin.setRange(10, 5000)
        self.undo_limit_spin.setSingleStep(10)
        self.undo_limit_spin.setValue(max(10, int(undo_limit)))
        self.undo_limit_spin.setSuffix(" 回")

        form_layout = QFormLayout()
        form_layout.addRow("アンドゥ履歴数", self.undo_limit_spin)

        shortcut_values = shortcuts or {}
        shortcut_group = self._build_shortcut_group(shortcut_values)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form_layout)
        layout.addWidget(shortcut_group)
        layout.addWidget(buttons)

    def _build_shortcut_group(self, shortcut_values: dict[str, list[str]]) -> QGroupBox:
        group = QGroupBox("ショートカットキー", self)
        group_layout = QVBoxLayout(group)

        form = QFormLayout()
        # 各操作につき「主」「副」の 2 つのキー割り当てを持てるようにする。
        self._shortcut_edits: dict[str, tuple[QKeySequenceEdit, QKeySequenceEdit]] = {}
        for spec in SHORTCUT_SPECS:
            sequences = shortcut_values.get(spec.key, list(spec.defaults))
            primary = QKeySequenceEdit(group)
            secondary = QKeySequenceEdit(group)
            if len(sequences) >= 1:
                primary.setKeySequence(QKeySequence(sequences[0]))
            if len(sequences) >= 2:
                secondary.setKeySequence(QKeySequence(sequences[1]))

            row_widget = QWidget(group)
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.addWidget(primary)
            row_layout.addWidget(secondary)
            form.addRow(spec.label, row_widget)
            self._shortcut_edits[spec.key] = (primary, secondary)

        group_layout.addLayout(form)

        reset_button = QPushButton("ショートカットを既定に戻す", group)
        reset_button.clicked.connect(self._reset_shortcuts_to_defaults)
        group_layout.addWidget(reset_button)

        return group

    def _reset_shortcuts_to_defaults(self) -> None:
        for spec in SHORTCUT_SPECS:
            primary, secondary = self._shortcut_edits[spec.key]
            defaults = list(spec.defaults)
            primary.setKeySequence(QKeySequence(defaults[0]) if len(defaults) >= 1 else QKeySequence())
            secondary.setKeySequence(QKeySequence(defaults[1]) if len(defaults) >= 2 else QKeySequence())

    def undo_limit(self) -> int:
        return self.undo_limit_spin.value()

    def shortcuts(self) -> dict[str, list[str]]:
        result: dict[str, list[str]] = {}
        for key, (primary, secondary) in self._shortcut_edits.items():
            sequences: list[str] = []
            for edit in (primary, secondary):
                sequence = edit.keySequence()
                if sequence.isEmpty():
                    continue
                text = sequence.toString(QKeySequence.SequenceFormat.PortableText)
                if text and text not in sequences:
                    sequences.append(text)
            result[key] = sequences
        return result

    def accept(self) -> None:
        # 別々の操作へ同じキーが割り当てられていないかを検証する。
        assigned: dict[str, str] = {}
        for key, sequences in self.shortcuts().items():
            for sequence in sequences:
                if sequence in assigned and assigned[sequence] != key:
                    QMessageBox.warning(
                        self,
                        "ショートカットの重複",
                        f"「{sequence}」が複数の操作に割り当てられています。\n"
                        "重複しないように設定し直してください。",
                    )
                    return
                assigned[sequence] = key
        super().accept()
