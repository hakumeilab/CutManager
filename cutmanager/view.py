from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QDate, QEvent, QModelIndex, QPoint, QRect, QSize, QTimer, Qt, Signal
from PySide6.QtGui import (
    QColor,
    QIcon,
    QKeySequence,
    QLinearGradient,
    QPainter,
    QPainterPath,
    QPalette,
    QPen,
    QPixmap,
)
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemDelegate,
    QAbstractItemView,
    QCalendarWidget,
    QComboBox,
    QDateEdit,
    QHeaderView,
    QLineEdit,
    QMenu,
    QStyledItemDelegate,
    QStyle,
    QStyleOptionButton,
    QStyleOptionViewItem,
    QTableView,
    QWidget,
    QWidgetAction,
)

from .constants import (
    BG_LOAD_COUNT_OPTIONS,
    COLUMN_BG_DATE,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_DELIVERY_DATE,
    COLUMN_STATUS,
    COLUMN_THUMBNAIL,
    COLUMN_TP_DATE,
    COLUMN_TP_LOAD_COUNT,
    COLUMN_VIDEO_PATH,
    STATUS_OPTIONS,
    TP_LOAD_COUNT_OPTIONS,
)


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


class CandidateEditorComboBox(QComboBox):
    confirmRequested = Signal()

    def __init__(self, parent=None, *, editable: bool = False) -> None:
        super().__init__(parent)
        self.setAutoFillBackground(True)
        self.setEditable(editable)

    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.confirmRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class CalendarEditor(QDateEdit):
    confirmRequested = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setCalendarPopup(True)
        self.setDisplayFormat("yyyy/MM/dd")
        self.setAutoFillBackground(True)
        self._calendar_menu: QMenu | None = None

    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.confirmRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)

    def open_calendar_menu(self) -> None:
        menu = QMenu(self)
        menu.setObjectName("calendarMenu")
        calendar = QCalendarWidget(menu)
        calendar.setObjectName("calendarPopup")
        calendar.setSelectedDate(self.date())
        calendar.setGridVisible(True)
        # ポップアップ内で日セルが潰れ、2 桁の日付が見切れないよう十分な領域を確保する。
        calendar.setMinimumSize(300, 260)
        calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        date_view = calendar.findChild(QAbstractItemView)
        if date_view is not None:
            date_view.setMinimumSize(300, 210)
        calendar.clicked.connect(lambda date, menu=menu: self._select_calendar_date(date, menu))
        calendar.activated.connect(lambda date, menu=menu: self._select_calendar_date(date, menu))
        menu.setStyleSheet(
            "QMenu#calendarMenu { border: 1px solid #94a3b8; border-radius: 6px; padding: 6px; }"
            "QCalendarWidget QWidget { alternate-background-color: #f8fbff; }"
            "QCalendarWidget QToolButton { border: 0px; border-radius: 4px; padding: 4px 8px; font-weight: 600; }"
            "QCalendarWidget QToolButton:hover { background: #dbeafe; }"
            "QCalendarWidget QAbstractItemView { border: 0px; selection-background-color: #2563eb; selection-color: #eff6ff; }"
        )
        action = QWidgetAction(menu)
        action.setDefaultWidget(calendar)
        menu.addAction(action)
        self._calendar_menu = menu
        menu.popup(self.mapToGlobal(QPoint(0, self.height())))

    def _select_calendar_date(self, date: QDate, menu: QMenu) -> None:
        self.setDate(date)
        menu.close()
        self.confirmRequested.emit()


class CutItemDelegate(QStyledItemDelegate):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._active_editor: QWidget | None = None
        # 縮小済みサムネイルのキャッシュ（描画のたびの再スケールを避ける）。
        self._scaled_thumbnail_cache: dict[tuple[int, int, int], QPixmap] = {}
        # 生成中に表示するスケルトンのライトスイープ用アニメーション。
        self._skeleton_phase = 0.0
        self._skeleton_timer = QTimer(self)
        self._skeleton_timer.setInterval(33)  # 約 30fps
        self._skeleton_timer.timeout.connect(self._advance_skeleton)

    def set_skeleton_animating(self, active: bool) -> None:
        if active:
            if not self._skeleton_timer.isActive():
                self._skeleton_timer.start()
        else:
            if self._skeleton_timer.isActive():
                self._skeleton_timer.stop()
            self._update_thumbnail_column()

    def _advance_skeleton(self) -> None:
        self._skeleton_phase = (self._skeleton_phase + 0.06) % 1.0
        self._update_thumbnail_column()

    def _update_thumbnail_column(self) -> None:
        # サムネイル列だけを再描画してアニメーション負荷を抑える。
        view = self.parent()
        if not isinstance(view, QTableView):
            return
        header = view.horizontalHeader()
        if header is None or view.isColumnHidden(COLUMN_THUMBNAIL):
            return
        x = header.sectionViewportPosition(COLUMN_THUMBNAIL)
        width = header.sectionSize(COLUMN_THUMBNAIL)
        viewport = view.viewport()
        viewport.update(x, 0, width, viewport.height())

    def createEditor(self, parent, option, index):
        if self._candidate_options(index.column()) is not None:
            options = self._candidate_options(index.column()) or ()
            editor = CandidateEditorComboBox(parent, editable=index.column() != COLUMN_STATUS)
            editor.addItems(options)
            editor.confirmRequested.connect(lambda: self._commit_and_close(editor, move_down=True))
            editor.activated.connect(lambda *_args: self._commit_and_close(editor, move_down=True))
            QTimer.singleShot(0, editor.showPopup)
        elif self._is_date_column(index.column()):
            editor = CalendarEditor(parent)
            editor.confirmRequested.connect(lambda: self._commit_and_close(editor, move_down=True))
            QTimer.singleShot(0, editor.open_calendar_menu)
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
            if combo_index >= 0:
                editor.setCurrentIndex(combo_index)
            elif editor.isEditable():
                editor.setEditText(value)
            else:
                editor.setCurrentIndex(0)
            return
        if isinstance(editor, QDateEdit):
            value = str(index.data(Qt.ItemDataRole.EditRole) or "").strip()
            date = QDate.fromString(value, "yyyy/MM/dd")
            editor.setDate(date if date.isValid() else QDate.currentDate())
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index) -> None:
        value = self._editor_value(editor)
        if value is None:
            super().setModelData(editor, model, index)
            return

        # 複数セルを選択した状態での編集は、同じ列の選択セルすべてへ同じ値を入れる。
        fill_targets = self._fill_targets(model, index)
        if not fill_targets:
            model.setData(index, value, Qt.ItemDataRole.EditRole)
            return
        self._commit_with_fill(model, index, value, fill_targets)

    @staticmethod
    def _editor_value(editor) -> str | None:
        if isinstance(editor, QComboBox):
            return editor.currentText()
        if isinstance(editor, QDateEdit):
            return editor.date().toString("yyyy/MM/dd")
        if isinstance(editor, QLineEdit):
            return editor.text()
        return None

    def _fill_targets(self, model, index) -> list[QModelIndex]:
        view = self.parent()
        if not isinstance(view, QAbstractItemView):
            return []
        selection_model = view.selectionModel()
        if selection_model is None:
            return []
        selected = selection_model.selectedIndexes()
        if len(selected) <= 1:
            return []

        column = index.column()
        targets: list[QModelIndex] = []
        for other in selected:
            if other.column() != column:
                continue
            if other.row() == index.row():
                continue
            if not bool(other.flags() & Qt.ItemFlag.ItemIsEditable):
                continue
            targets.append(QModelIndex(other))
        return targets

    def _commit_with_fill(self, model, index, value: str, targets: list[QModelIndex]) -> None:
        source_model = getattr(model, "sourceModel", lambda: None)()
        map_to_source = getattr(model, "mapToSource", None)
        apply_changes = getattr(source_model, "apply_cell_changes", None)
        if source_model is None or map_to_source is None or apply_changes is None:
            # プロキシ/一括APIが無い場合は 1 セルずつ書き込む。
            model.setData(index, value, Qt.ItemDataRole.EditRole)
            for target in targets:
                model.setData(target, value, Qt.ItemDataRole.EditRole)
            return

        actual_rows = source_model.actual_row_count()
        changes: list[tuple[int, int, str]] = []
        primary_is_virtual = False
        for proxy_index in [index, *targets]:
            source_index = map_to_source(proxy_index)
            if source_index.row() >= actual_rows:
                # 末尾の仮想（空）行はここでは扱えないため個別に委譲する。
                if proxy_index is index:
                    primary_is_virtual = True
                continue
            changes.append((source_index.row(), source_index.column(), value))

        if primary_is_virtual:
            model.setData(index, value, Qt.ItemDataRole.EditRole)
        if changes:
            # 1 回の履歴コマンドにまとめ、まとめてアンドゥできるようにする。
            apply_changes(changes)

    def paint(self, painter, option, index) -> None:
        if index.column() == COLUMN_THUMBNAIL:
            self._paint_thumbnail(painter, option, index)
            return

        option_copy = type(option)(option)
        self.initStyleOption(option_copy, index)
        indicator_option = type(option_copy)(option_copy)

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

        if self._is_icon_column(index.column()) and self._is_editing_index(index):
            option_copy.text = ""
        elif self._is_icon_column(index.column()):
            option_copy.rect = option_copy.rect.adjusted(0, 0, -18, 0)
        super().paint(painter, option_copy, index)
        if self._is_candidate_column(index.column()) and not self._is_editing_index(index):
            self._paint_candidate_indicator(painter, indicator_option)
        elif self._is_date_column(index.column()) and not self._is_editing_index(index):
            self._paint_calendar_indicator(painter, indicator_option)

    def sizeHint(self, option, index) -> QSize:
        # 元サムネイルが大きくても行高/列幅の推定を膨らませない（セル内に収める描画のため）。
        if index.column() == COLUMN_THUMBNAIL:
            return QSize(48, 24)
        return super().sizeHint(option, index)

    def _paint_thumbnail(self, painter, option, index) -> None:
        opt = type(option)(option)
        self.initStyleOption(opt, index)
        # 背景・選択は標準スタイルで描画し、テキスト/アイコンは自前で扱う。
        opt.text = ""
        opt.icon = QIcon()
        opt.features &= ~QStyleOptionViewItem.ViewItemFeature.HasDecoration

        widget = opt.widget
        style = widget.style() if widget is not None else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, widget)

        margin = 2
        target = opt.rect.adjusted(margin, margin, -margin, -margin)
        if target.width() <= 0 or target.height() <= 0:
            return

        pixmap = index.data(Qt.ItemDataRole.DecorationRole)
        if not isinstance(pixmap, QPixmap) or pixmap.isNull():
            # サムネイル生成待ちの行はスケルトン＋ライトスイープを表示する。
            video_path = str(
                index.siblingAtColumn(COLUMN_VIDEO_PATH).data(Qt.ItemDataRole.DisplayRole) or ""
            )
            if video_path.strip():
                self._paint_skeleton(painter, target, opt.palette)
            return

        scaled = self._scaled_thumbnail(pixmap, target.width(), target.height())
        x = target.x() + (target.width() - scaled.width()) // 2
        y = target.y() + (target.height() - scaled.height()) // 2
        painter.drawPixmap(x, y, scaled)

    def _paint_skeleton(self, painter, rect, palette) -> None:
        radius = 3
        base_color = palette.color(QPalette.ColorRole.Base)
        dark = (0.299 * base_color.red() + 0.587 * base_color.green() + 0.114 * base_color.blue()) < 128
        if dark:
            fill = QColor("#243044")
            sweep = QColor(255, 255, 255, 45)
        else:
            fill = QColor("#e2e8f0")
            sweep = QColor(255, 255, 255, 200)

        path = QPainterPath()
        path.addRoundedRect(rect.x(), rect.y(), rect.width(), rect.height(), radius, radius)

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.fillPath(path, fill)

        # 左から右へ移動する光の帯。
        band = max(12, rect.width() // 3)
        travel = rect.width() + band
        center = rect.x() - band / 2 + self._skeleton_phase * travel
        gradient = QLinearGradient(center - band / 2, 0, center + band / 2, 0)
        transparent = QColor(sweep)
        transparent.setAlpha(0)
        gradient.setColorAt(0.0, transparent)
        gradient.setColorAt(0.5, sweep)
        gradient.setColorAt(1.0, transparent)
        painter.setClipPath(path)
        painter.fillRect(rect, gradient)
        painter.restore()

    def _scaled_thumbnail(self, pixmap: QPixmap, width: int, height: int) -> QPixmap:
        key = (pixmap.cacheKey(), width, height)
        cached = self._scaled_thumbnail_cache.get(key)
        if cached is not None:
            return cached
        scaled = pixmap.scaled(
            QSize(width, height),
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        # キャッシュの肥大化を防ぐ（列/行サイズやサムネ更新の組み合わせ上限）。
        if len(self._scaled_thumbnail_cache) > 256:
            self._scaled_thumbnail_cache.clear()
        self._scaled_thumbnail_cache[key] = scaled
        return scaled

    def editorEvent(self, event, model, option, index) -> bool:
        if (
            event.type() == QEvent.Type.MouseButtonRelease
            and event.button() == Qt.MouseButton.LeftButton
            and self._is_icon_column(index.column())
            and self._indicator_rect(option.rect).contains(event.pos())
        ):
            view = self.parent()
            index_copy = QModelIndex(index)
            if hasattr(view, "edit"):
                QTimer.singleShot(0, lambda view=view, index_copy=index_copy: view.edit(index_copy))
                return True
        return super().editorEvent(event, model, option, index)

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
    def _candidate_options(column: int) -> tuple[str, ...] | None:
        if column == COLUMN_STATUS:
            return STATUS_OPTIONS
        if column == COLUMN_TP_LOAD_COUNT:
            return TP_LOAD_COUNT_OPTIONS
        if column == COLUMN_BG_LOAD_COUNT:
            return BG_LOAD_COUNT_OPTIONS
        return None

    @classmethod
    def _is_candidate_column(cls, column: int) -> bool:
        return cls._candidate_options(column) is not None

    @staticmethod
    def _is_date_column(column: int) -> bool:
        return column in (COLUMN_TP_DATE, COLUMN_BG_DATE, COLUMN_DELIVERY_DATE)

    @classmethod
    def _is_icon_column(cls, column: int) -> bool:
        return cls._is_candidate_column(column) or cls._is_date_column(column)

    @staticmethod
    def _paint_candidate_indicator(painter, option) -> None:
        rect = option.rect
        if rect.width() < 18 or rect.height() < 14:
            return

        palette = option.palette
        icon_color = palette.color(QPalette.ColorRole.Mid)
        center = CutItemDelegate._indicator_rect(rect).center()

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        pen = QPen(icon_color)
        pen.setWidth(1)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.drawLine(center + QPoint(-3, -2), center + QPoint(0, 1))
        painter.drawLine(center + QPoint(0, 1), center + QPoint(3, -2))
        painter.restore()

    @staticmethod
    def _paint_calendar_indicator(painter, option) -> None:
        rect = option.rect
        if rect.width() < 18 or rect.height() < 14:
            return

        icon_rect = CutItemDelegate._indicator_rect(rect).adjusted(2, 2, -2, -2)
        icon_color = option.palette.color(QPalette.ColorRole.Mid)

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        pen = QPen(icon_color)
        pen.setWidth(1)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(icon_rect.adjusted(0, 2, -1, -1), 2, 2)
        painter.drawLine(icon_rect.left(), icon_rect.top() + 5, icon_rect.right() - 1, icon_rect.top() + 5)
        painter.drawLine(icon_rect.left() + 3, icon_rect.top(), icon_rect.left() + 3, icon_rect.top() + 3)
        painter.drawLine(icon_rect.right() - 4, icon_rect.top(), icon_rect.right() - 4, icon_rect.top() + 3)
        painter.restore()

    @staticmethod
    def _indicator_rect(cell_rect: QRect) -> QRect:
        return QRect(
            cell_rect.right() - 17,
            cell_rect.center().y() - 8,
            16,
            16,
        )

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
