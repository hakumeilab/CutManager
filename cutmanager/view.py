from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QEvent, QRect, QTimer, Qt, Signal
from PySide6.QtGui import QColor, QKeySequence, QPainter, QPainterPath, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemDelegate,
    QAbstractItemView,
    QComboBox,
    QHeaderView,
    QLineEdit,
    QStyledItemDelegate,
    QStyle,
    QStyleOptionButton,
    QTableView,
    QWidget,
)

from .constants import COLUMN_STATUS, STATUS_OPTIONS


class CellEditorLineEdit(QLineEdit):
    confirmRequested = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFrame(False)
        self.setTextMargins(2, 0, 2, 0)
        self.setStyleSheet("QLineEdit { border: 0px; padding: 0px 2px; margin: 0px; }")

    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.confirmRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class StatusEditorComboBox(QComboBox):
    confirmRequested = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setAutoFillBackground(True)

    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.confirmRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class CutItemDelegate(QStyledItemDelegate):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._active_editor: QWidget | None = None

    def createEditor(self, parent, option, index):
        if index.column() == COLUMN_STATUS:
            editor = StatusEditorComboBox(parent)
            editor.addItems(STATUS_OPTIONS)
            editor.confirmRequested.connect(lambda: self._commit_and_close(editor, move_down=True))
            editor.activated.connect(lambda *_args: self._commit_and_close(editor, move_down=True))
            QTimer.singleShot(0, editor.showPopup)
        else:
            editor = CellEditorLineEdit(parent)
            editor.confirmRequested.connect(lambda: self._commit_and_close(editor, move_down=True))

        self._active_editor = editor
        editor.setProperty("_cutmanager_row", index.row())
        editor.setProperty("_cutmanager_column", index.column())
        editor.destroyed.connect(self._clear_active_editor)
        return editor

    def setEditorData(self, editor, index) -> None:
        if isinstance(editor, QComboBox):
            value = str(index.data(Qt.ItemDataRole.EditRole) or "")
            combo_index = editor.findText(value)
            editor.setCurrentIndex(combo_index if combo_index >= 0 else 0)
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index) -> None:
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)
            return
        super().setModelData(editor, model, index)

    def paint(self, painter, option, index) -> None:
        option_copy = type(option)(option)
        self.initStyleOption(option_copy, index)

        background = index.data(Qt.ItemDataRole.BackgroundRole)
        if isinstance(background, QColor):
            fill_color = QColor(background)
            if option_copy.state & QStyle.StateFlag.State_Selected:
                highlight = option_copy.palette.color(QPalette.ColorRole.Highlight)
                fill_color = self._blend_colors(fill_color, highlight, 0.16)
            painter.fillRect(option_copy.rect, fill_color)
            option_copy.backgroundBrush = QColor(fill_color)

        foreground = index.data(Qt.ItemDataRole.ForegroundRole)
        if isinstance(foreground, QColor):
            option_copy.palette.setColor(QPalette.ColorRole.Text, foreground)
            option_copy.palette.setColor(QPalette.ColorRole.WindowText, foreground)
            option_copy.palette.setColor(QPalette.ColorRole.HighlightedText, foreground)

        option_copy.state &= ~QStyle.StateFlag.State_Selected
        option_copy.state &= ~QStyle.StateFlag.State_HasFocus

        if index.column() == COLUMN_STATUS and self._is_editing_index(index):
            option_copy.text = ""
        super().paint(painter, option_copy, index)

    def current_editor(self) -> QWidget | None:
        return self._active_editor

    def _commit_and_close(self, editor: QWidget, *, move_down: bool) -> None:
        row_value = editor.property("_cutmanager_row")
        column_value = editor.property("_cutmanager_column")
        row = -1 if row_value is None else int(row_value)
        column = -1 if column_value is None else int(column_value)
        self.commitData.emit(editor)
        self.closeEditor.emit(editor, QAbstractItemDelegate.EndEditHint.NoHint)
        parent_view = self.parent()
        if move_down and row >= 0 and column >= 0 and hasattr(parent_view, "_move_to_cell_below"):
            QTimer.singleShot(10, lambda row=row, column=column, view=parent_view: view._move_to_cell_below(row, column))

    def _clear_active_editor(self, *_args) -> None:
        self._active_editor = None

    def _is_editing_index(self, index) -> bool:
        if self._active_editor is None:
            return False
        row_value = self._active_editor.property("_cutmanager_row")
        column_value = self._active_editor.property("_cutmanager_column")
        if row_value is None or column_value is None:
            return False
        return int(row_value) == index.row() and int(column_value) == index.column()

    @staticmethod
    def _blend_colors(base: QColor, overlay: QColor, overlay_alpha: float) -> QColor:
        alpha = max(0.0, min(1.0, overlay_alpha))
        inverse = 1.0 - alpha
        return QColor(
            round((base.red() * inverse) + (overlay.red() * alpha)),
            round((base.green() * inverse) + (overlay.green() * alpha)),
            round((base.blue() * inverse) + (overlay.blue() * alpha)),
        )


class FilterHeaderView(QHeaderView):
    filterButtonClicked = Signal(int)

    BUTTON_SIZE = 18
    BUTTON_MARGIN = 4

    def __init__(self, orientation: Qt.Orientation, parent=None) -> None:
        super().__init__(orientation, parent)
        self._filtered_columns: set[int] = set()

    def set_filtered_columns(self, columns: set[int]) -> None:
        self._filtered_columns = set(columns)
        self.viewport().update()

    def paintSection(self, painter, rect, logicalIndex) -> None:
        super().paintSection(painter, rect, logicalIndex)
        if not rect.isValid() or rect.width() <= self.BUTTON_SIZE + (self.BUTTON_MARGIN * 2):
            return

        if self._is_dark_palette():
            button = QColor("#172033")
            mid = QColor("#334155")
            highlighted_text = QColor("#eff6ff")
            accent = QColor("#3b82f6")
            icon_idle = QColor("#9fb0c9")
        else:
            button = QColor("#ffffff")
            mid = QColor("#cbd5e1")
            highlighted_text = QColor("#eff6ff")
            accent = QColor("#2563eb")
            icon_idle = QColor("#475569")

        button_rect = self._button_rect(rect)
        option = QStyleOptionButton()
        option.rect = button_rect
        option.state = QStyle.StateFlag.State_Enabled
        if logicalIndex in self._filtered_columns:
            option.state |= QStyle.StateFlag.State_On

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(mid)
        painter.setBrush(button if logicalIndex not in self._filtered_columns else QColor(accent))
        painter.drawRoundedRect(button_rect.adjusted(0, 0, -1, -1), 6, 6)
        painter.restore()

        icon_color = QColor(highlighted_text) if logicalIndex in self._filtered_columns else icon_idle
        icon_rect = button_rect.adjusted(5, 4, -5, -4)
        top_y = icon_rect.top()
        mid_y = icon_rect.center().y() - 1
        bottom_y = icon_rect.bottom()
        center_x = icon_rect.center().x()

        funnel_path = QPainterPath()
        funnel_path.moveTo(icon_rect.left(), top_y)
        funnel_path.lineTo(icon_rect.right(), top_y)
        funnel_path.lineTo(center_x + 2, mid_y)
        funnel_path.lineTo(center_x + 2, bottom_y)
        funnel_path.lineTo(center_x - 2, bottom_y)
        funnel_path.lineTo(center_x - 2, mid_y)
        funnel_path.closeSubpath()

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(icon_color)
        painter.drawPath(funnel_path)
        painter.restore()

    def mousePressEvent(self, event) -> None:
        logical_index = self.logicalIndexAt(event.pos())
        if logical_index >= 0:
            section_rect = self._section_rect(logical_index)
            if self._button_rect(section_rect).contains(event.pos()):
                self.filterButtonClicked.emit(logical_index)
                event.accept()
                return
        super().mousePressEvent(event)

    def _section_rect(self, logical_index: int) -> QRect:
        return QRect(
            self.sectionViewportPosition(logical_index),
            0,
            self.sectionSize(logical_index),
            self.height(),
        )

    def _button_rect(self, section_rect: QRect) -> QRect:
        return QRect(
            section_rect.right() - self.BUTTON_SIZE - self.BUTTON_MARGIN,
            section_rect.center().y() - (self.BUTTON_SIZE // 2),
            self.BUTTON_SIZE,
            self.BUTTON_SIZE,
        )

    def _is_dark_palette(self) -> bool:
        app = QApplication.instance()
        if app is not None:
            try:
                scheme = app.styleHints().colorScheme()
                if scheme == Qt.ColorScheme.Dark:
                    return True
                if scheme == Qt.ColorScheme.Light:
                    return False
            except AttributeError:
                pass
        base = self.palette().color(QPalette.ColorRole.Base)
        luminance = (0.299 * base.red()) + (0.587 * base.green()) + (0.114 * base.blue())
        return luminance < 128


class CutTableView(QTableView):
    clearRequested = Signal()
    addRowRequested = Signal()
    deleteRowsRequested = Signal()
    copyRequested = Signal()
    pasteRequested = Signal()
    pathsDropped = Signal(list)
    dragStateChanged = Signal(bool)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.viewport().installEventFilter(self)
        self.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)

    def eventFilter(self, source, event) -> bool:
        if source is self.viewport():
            if event.type() in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
                if self._has_local_paths(event.mimeData()):
                    self.dragStateChanged.emit(True)
                    event.acceptProposedAction()
                    return True

            if event.type() == QEvent.Type.DragLeave:
                self.dragStateChanged.emit(False)

            if event.type() == QEvent.Type.Drop:
                paths = self._extract_local_paths(event.mimeData())
                self.dragStateChanged.emit(False)
                if paths:
                    self.pathsDropped.emit(paths)
                    event.acceptProposedAction()
                    return True

        return super().eventFilter(source, event)

    def keyPressEvent(self, event) -> None:
        modifiers = event.modifiers()

        if self.state() != QAbstractItemView.State.EditingState:
            if event.matches(QKeySequence.StandardKey.Copy):
                self.copyRequested.emit()
                event.accept()
                return

            if event.matches(QKeySequence.StandardKey.Paste):
                self.pasteRequested.emit()
                event.accept()
                return

        if event.key() == Qt.Key.Key_Delete and modifiers == Qt.KeyboardModifier.ControlModifier:
            self.deleteRowsRequested.emit()
            event.accept()
            return

        if event.key() == Qt.Key.Key_Delete and self.state() != QAbstractItemView.State.EditingState:
            self.clearRequested.emit()
            event.accept()
            return

        if event.key() == Qt.Key.Key_Insert and self.state() != QAbstractItemView.State.EditingState:
            self.addRowRequested.emit()
            event.accept()
            return

        direct_input_text = self._direct_input_text(event)
        if self.state() != QAbstractItemView.State.EditingState and direct_input_text:
            current = self.currentIndex()
            if current.isValid():
                self.edit(current)
                QTimer.singleShot(0, lambda text=direct_input_text: self._apply_initial_text(text))
                event.accept()
                return

        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and self.state() != QAbstractItemView.State.EditingState:
            current = self.currentIndex()
            if current.isValid():
                self.edit(current)
                event.accept()
                return

        super().keyPressEvent(event)

    @staticmethod
    def _has_local_paths(mime_data) -> bool:
        return bool(CutTableView._extract_local_paths(mime_data))

    @staticmethod
    def _extract_local_paths(mime_data) -> list[Path]:
        if mime_data is None or not mime_data.hasUrls():
            return []

        paths: list[Path] = []
        for url in mime_data.urls():
            if not url.isLocalFile():
                continue
            local_path = url.toLocalFile()
            if local_path:
                paths.append(Path(local_path))
        return paths

    @staticmethod
    def _direct_input_text(event) -> str:
        text = event.text()
        blocked_modifiers = (
            Qt.KeyboardModifier.ControlModifier
            | Qt.KeyboardModifier.AltModifier
            | Qt.KeyboardModifier.MetaModifier
        )
        if event.modifiers() & blocked_modifiers:
            return ""

        if event.key() in (
            Qt.Key.Key_Return,
            Qt.Key.Key_Enter,
            Qt.Key.Key_Tab,
            Qt.Key.Key_Backtab,
            Qt.Key.Key_Escape,
        ):
            return ""

        if text and text.isprintable() and not text.isspace():
            return text

        if Qt.Key.Key_0 <= event.key() <= Qt.Key.Key_9:
            return str(event.key() - Qt.Key.Key_0)

        return ""

    def _apply_initial_text(self, text: str) -> None:
        editor = None
        delegate = self.itemDelegate()
        if isinstance(delegate, CutItemDelegate):
            editor = delegate.current_editor()
        if not isinstance(editor, QLineEdit):
            editor = QApplication.focusWidget()
        if not isinstance(editor, QLineEdit):
            return

        editor.selectAll()
        editor.insert(text)

    def _move_to_cell_below(self, row: int, column: int) -> None:
        model = self.model()
        if model is None:
            return

        next_index = model.index(row + 1, column)
        if not next_index.isValid():
            return

        self.setCurrentIndex(next_index)
        self.scrollTo(next_index)
