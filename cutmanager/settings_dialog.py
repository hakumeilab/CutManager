from __future__ import annotations

from PySide6.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QSpinBox, QVBoxLayout


class SettingsDialog(QDialog):
    def __init__(self, undo_limit: int, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("環境設定")

        self.undo_limit_spin = QSpinBox(self)
        self.undo_limit_spin.setRange(10, 5000)
        self.undo_limit_spin.setSingleStep(10)
        self.undo_limit_spin.setValue(max(10, int(undo_limit)))
        self.undo_limit_spin.setSuffix(" 回")

        form_layout = QFormLayout()
        form_layout.addRow("アンドゥ履歴数", self.undo_limit_spin)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form_layout)
        layout.addWidget(buttons)

    def undo_limit(self) -> int:
        return self.undo_limit_spin.value()
