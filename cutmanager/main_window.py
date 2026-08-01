from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import (
    QDate,
    QDir,
    QEvent,
    QPoint,
    QProcess,
    QSettings,
    QThread,
    QTimer,
    Qt,
    QUrl,
    Signal,
)
from PySide6.QtGui import (
    QAction,
    QCloseEvent,
    QColor,
    QDesktopServices,
    QDragEnterEvent,
    QDropEvent,
    QPainter,
    QPalette,
    QPen,
)
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QProgressDialog,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from .constants import (
    BG_FILE_EXTENSIONS,
    COLUMN_BG_DATE,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_CUT_NUMBER,
    COLUMN_DELIVERY_DATE,
    COLUMN_THUMBNAIL,
    COLUMN_VIDEO_PATH,
    CSV_FILE_FILTER,
    CSV_HEADERS,
    IMPORT_DATE_FORMAT,
    COLUMN_STATUS,
    COLUMN_TP_DATE,
    COLUMN_TP_LOAD_COUNT,
    PROJECT_FILE_EXTENSION,
    PROJECT_SAVE_FILTER,
    SUPPORTED_PROJECT_EXTENSIONS,
    VIDEO_FILE_EXTENSIONS,
    WINDOW_SIZE,
    WINDOW_TITLE,
)
from .csv_io import CsvLoadError, load_csv_file, save_csv_file
from .filter_popup import ColumnFilterPopup
from .folder_import import apply_material_updates, build_rows_from_dropped_folders
from .history import HistoryManager
from .model import CutTableModel, NON_DATA_COLUMNS
from .proxy import CutFilterProxyModel
from .thumbnails import ThumbnailProvider
from .settings_dialog import SettingsDialog
from .shortcuts import ShortcutManager
from .update_manager import (
    RELEASES_PAGE_URL,
    PreparedUpdate,
    UpdateAsset,
    UpdateCheckResult,
    UpdateCheckWorker,
    UpdateDownloadWorker,
    UpdateError,
    human_readable_size,
    prepare_update,
)
from .video_import import apply_videos_to_rows, build_rows_from_video_files
from .view import CutItemDelegate, CutTableView, FilterHeaderView


def calculate_cut_summary(rows: list[list[str]]) -> dict[str, int]:
    total_cuts = 0
    delivered = 0
    total_tp = 0
    tp_done = 0
    total_bg = 0
    bg_done = 0
    remaining_tp = 0
    remaining_bg = 0
    remaining_delivery = 0
    shared = 0
    bank = 0
    missing = 0

    for row in rows:
        normalized_row = _normalize_summary_row(row)
        status = normalized_row[COLUMN_STATUS].strip()
        if status == "欠番":
            missing += 1
            continue

        total_cuts += 1
        if status == "兼用":
            shared += 1
        if status == "BANK":
            bank += 1
            continue

        if normalized_row[COLUMN_DELIVERY_DATE].strip():
            delivered += 1
        else:
            remaining_delivery += 1

        tp_value = normalized_row[COLUMN_TP_LOAD_COUNT].strip()
        bg_value = normalized_row[COLUMN_BG_LOAD_COUNT].strip()
        tp_required = tp_value != "BGOnly"
        bg_required = bg_value != "全セル"

        if tp_required:
            total_tp += 1
        if bg_required:
            total_bg += 1

        if tp_required and tp_value:
            tp_done += 1
        elif tp_required:
            remaining_tp += 1

        if bg_required and bg_value:
            bg_done += 1
        elif bg_required:
            remaining_bg += 1

    return {
        "total_cuts": total_cuts,
        "delivered": delivered,
        "remaining_delivery": remaining_delivery,
        "total_tp": total_tp,
        "tp_done": tp_done,
        "remaining_tp": remaining_tp,
        "total_bg": total_bg,
        "bg_done": bg_done,
        "remaining_bg": remaining_bg,
        "shared": shared,
        "bank": bank,
        "missing": missing,
    }


def _normalize_summary_row(row: list[str]) -> list[str]:
    normalized = [""] * len(CSV_HEADERS)
    for index in range(min(len(row), len(CSV_HEADERS))):
        normalized[index] = "" if row[index] is None else str(row[index])
    return normalized


class RibbonToggleButton(QWidget):
    clicked = Signal(bool)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFixedSize(18, 18)
        self._checked = True

    def isChecked(self) -> bool:
        return self._checked

    def setChecked(self, checked: bool) -> None:
        if self._checked == checked:
            return
        self._checked = checked
        self.update()

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit(self._checked)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def paintEvent(self, event) -> None:
        palette = self.palette()
        icon_color = palette.color(QPalette.ColorRole.Mid)
        center = self.rect().center()
        y_offset = 1 if self._checked else -1

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        pen = QPen(icon_color)
        pen.setWidth(1)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        if self._checked:
            painter.drawLine(center + QPoint(-3, 1 + y_offset), center + QPoint(0, -2 + y_offset))
            painter.drawLine(center + QPoint(0, -2 + y_offset), center + QPoint(3, 1 + y_offset))
        else:
            painter.drawLine(center + QPoint(-3, -2 + y_offset), center + QPoint(0, 1 + y_offset))
            painter.drawLine(center + QPoint(0, 1 + y_offset), center + QPoint(3, -2 + y_offset))


class RibbonToggleRow(QWidget):
    clicked = Signal()

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit()
            event.accept()
            return
        super().mouseReleaseEvent(event)


class MainWindow(QMainWindow):
    MAX_RECENT_FILES = 8
    LAST_SESSION_FILE_KEY = "lastSessionFile"
    UNDO_LIMIT_KEY = "undoLimit"
    EPISODE_MEMO_PREFIX = "episodeMemo/"
    HEADER_STATE_KEY = "tableHeaderState"
    COLUMN_VISIBILITY_KEY = "hiddenColumns"
    DEFAULT_UNDO_LIMIT = 100

    def __init__(self) -> None:
        super().__init__()

        self.current_file_path: str | None = None
        self.last_drop_summary = "-"
        self._sort_column = COLUMN_CUT_NUMBER
        self._sort_order = Qt.SortOrder.AscendingOrder
        self._pending_resort = False
        self._skip_close_confirmation = False
        self._drag_feedback_active = False
        self._drag_accept_cache: bool | None = None
        self._theme_apply_pending = False
        self._applying_theme_styles = False
        self._last_window_stylesheet = ""
        self._last_table_stylesheet = ""
        self._restoring_header_state = False
        self._syncing_section_size = False
        self.settings = QSettings("CutManager", "CutManager")
        self.shortcut_manager = ShortcutManager(self.settings)
        self.recent_files = self._load_recent_files()

        self.model = CutTableModel(parent=self)
        self.history = HistoryManager(self._load_undo_limit(), self)
        self.model.set_history_manager(self.history)
        self.thumbnail_provider = ThumbnailProvider(self)
        self.model.set_thumbnail_provider(self.thumbnail_provider)
        self.thumbnail_provider.thumbnailReady.connect(self.model.refresh_thumbnails_for_path)
        self.thumbnail_provider.progressChanged.connect(self._on_thumbnail_progress)
        self.proxy_model = CutFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.model)

        self.table_view = CutTableView(self)
        self.drop_hint_label = QLabel(self)
        self.file_path_label = QLabel(self)
        self.row_count_label = QLabel(self)
        self.modified_label = QLabel(self)
        self.drop_result_label = QLabel(self)
        self.summary_labels: dict[str, QLabel] = {}
        self.summary_ribbon_clip: QWidget | None = None
        self.summary_ribbon_body: QWidget | None = None
        self.summary_toggle_button: RibbonToggleButton | None = None
        self.summary_memo_edit: QPlainTextEdit | None = None
        self._updating_summary_memo = False
        self.drop_progress_bar = QProgressBar(self)
        self.thumbnail_progress_bar = QProgressBar(self)
        self.thumbnail_status_label = QLabel(self)
        self.file_menu = QMenu("ファイル", self)
        self.edit_menu = QMenu("編集", self)
        self.sort_menu = QMenu("並べ替え", self)
        self.view_menu = QMenu("表示", self)
        self.help_menu = QMenu("ヘルプ", self)
        self.column_visibility_actions: dict[int, QAction] = {}
        self.recent_files_menu = QMenu("最近開いたファイル", self)
        self._update_check_thread: QThread | None = None
        self._update_check_worker: UpdateCheckWorker | None = None
        self._update_download_thread: QThread | None = None
        self._update_download_worker: UpdateDownloadWorker | None = None
        self._download_progress_dialog: QProgressDialog | None = None

        self.new_action: QAction
        self.open_action: QAction
        self.save_action: QAction
        self.save_as_action: QAction
        self.undo_action: QAction
        self.redo_action: QAction
        self.copy_action: QAction
        self.paste_action: QAction
        self.add_row_action: QAction
        self.add_row_above_action: QAction
        self.add_row_below_action: QAction
        self.delete_row_action: QAction
        self.clear_values_action: QAction
        self.regenerate_thumbnails_action: QAction
        self.preferences_action: QAction
        self.restore_default_sort_action: QAction
        self.check_updates_action: QAction
        self.license_info_action: QAction

        self.setAcceptDrops(True)
        self.resize(*WINDOW_SIZE)

        self._create_actions()
        self._build_ui()
        self._connect_signals()
        self._connect_theme_signals()
        self._update_all_status()
        self._restore_last_session_file()

    def _create_actions(self) -> None:
        self.new_action = QAction("新規作成", self)
        self.new_action.triggered.connect(self.create_new_csv)

        self.open_action = QAction("開く", self)
        self.open_action.triggered.connect(self.open_csv_dialog)

        self.save_action = QAction("上書き保存", self)
        self.save_action.triggered.connect(self.save_csv)

        self.save_as_action = QAction("名前を付けて保存", self)
        self.save_as_action.triggered.connect(self.save_csv_as)

        self.undo_action = QAction("元に戻す", self)
        self.undo_action.setEnabled(False)
        self.undo_action.triggered.connect(self.undo)

        self.redo_action = QAction("やり直し", self)
        self.redo_action.setEnabled(False)
        self.redo_action.triggered.connect(self.redo)

        self.copy_action = QAction("コピー", self)
        self.copy_action.triggered.connect(self.copy_selected_cells)

        self.paste_action = QAction("貼り付け", self)
        self.paste_action.triggered.connect(self.paste_cells_from_clipboard)

        self.add_row_action = QAction("行追加", self)
        self.add_row_action.triggered.connect(lambda checked=False: self.add_row())

        self.add_row_above_action = QAction("上に行を追加", self)
        self.add_row_above_action.triggered.connect(self.add_row_above)

        self.add_row_below_action = QAction("下に行を追加", self)
        self.add_row_below_action.triggered.connect(self.add_row_below)

        self.delete_row_action = QAction("行削除", self)
        self.delete_row_action.triggered.connect(self.delete_selected_rows)

        self.clear_values_action = QAction("値を削除", self)
        self.clear_values_action.triggered.connect(self.clear_selected_cells)

        self.regenerate_thumbnails_action = QAction("サムネイル再生成", self)
        self.regenerate_thumbnails_action.setToolTip("すべての動画パスのサムネイルを作り直します。")
        self.regenerate_thumbnails_action.triggered.connect(self.regenerate_thumbnails)

        self.preferences_action = QAction("環境設定", self)
        self.preferences_action.triggered.connect(self.open_settings_dialog)

        self.restore_default_sort_action = QAction("カット番号順に戻す", self)
        self.restore_default_sort_action.setEnabled(False)
        self.restore_default_sort_action.triggered.connect(self._restore_default_sort)

        self.check_updates_action = QAction("更新を確認", self)
        self.check_updates_action.triggered.connect(self.check_for_updates)

        self.license_info_action = QAction("ライセンス情報", self)
        self.license_info_action.triggered.connect(self.show_license_info)

        for action in (
            self.new_action,
            self.open_action,
            self.save_action,
            self.save_as_action,
            self.undo_action,
            self.redo_action,
            self.copy_action,
            self.paste_action,
            self.add_row_action,
            self.add_row_above_action,
            self.add_row_below_action,
            self.delete_row_action,
            self.clear_values_action,
            self.regenerate_thumbnails_action,
            self.preferences_action,
            self.restore_default_sort_action,
            self.check_updates_action,
            self.license_info_action,
        ):
            self.addAction(action)

        # 環境設定で変更できるショートカットと対象アクションの対応表。
        self._shortcut_actions = {
            "new": self.new_action,
            "open": self.open_action,
            "save": self.save_action,
            "save_as": self.save_as_action,
            "undo": self.undo_action,
            "redo": self.redo_action,
            "copy": self.copy_action,
            "paste": self.paste_action,
            "add_row": self.add_row_action,
            "delete_row": self.delete_row_action,
        }
        self._apply_shortcuts()

    def _apply_shortcuts(self) -> None:
        for key, action in self._shortcut_actions.items():
            action.setShortcuts(self.shortcut_manager.key_sequences(key))

    def _build_ui(self) -> None:
        self.setWindowTitle(WINDOW_TITLE)

        self._build_menu_bar()
        self._build_menus()

        self.drop_hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_hint_label.setVisible(False)
        self.drop_hint_label.setText("ここに CSV / 素材フォルダー / PSD / 動画ファイル・Roll フォルダーをドロップ")

        self.table_view.setModel(self.proxy_model)
        self.cut_delegate = CutItemDelegate(self.table_view)
        self.table_view.setItemDelegate(self.cut_delegate)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        self.table_view.setTabKeyNavigation(True)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setWordWrap(False)
        self.table_view.verticalHeader().setDefaultSectionSize(24)
        self.table_view.verticalHeader().setMinimumWidth(44)

        header = FilterHeaderView(Qt.Orientation.Horizontal, self.table_view)
        self.table_view.setHorizontalHeader(header)
        header.setStretchLastSection(False)
        header.setSectionsClickable(True)
        header.setSectionsMovable(True)
        header.setSortIndicatorShown(True)
        header.setSortIndicator(self._sort_column, self._sort_order)
        header.sectionMoved.connect(self._save_header_state)
        self._apply_theme_styles()

        default_widths = [120, 180, 90, 110, 105, 115, 105, 115, 90, 105, 115, 80, 240, 110]
        for column, width in enumerate(default_widths):
            self.table_view.setColumnWidth(column, width)
        self._restore_header_state()
        self._apply_column_visibility()

        self._set_drag_feedback(False)

        container = QWidget(self)
        container.setObjectName("mainContainer")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self._build_summary_ribbon())
        layout.addWidget(self.table_view, 1)
        self.setCentralWidget(container)
        self.drop_hint_label.setParent(self.table_view.viewport())
        self.drop_hint_label.raise_()
        self._update_drop_hint_geometry()

        self.drop_progress_bar.setRange(0, 0)
        self.drop_progress_bar.setTextVisible(False)
        self.drop_progress_bar.setVisible(False)
        self.drop_progress_bar.setFixedWidth(160)

        self.thumbnail_progress_bar.setTextVisible(False)
        self.thumbnail_progress_bar.setVisible(False)
        self.thumbnail_progress_bar.setFixedWidth(140)
        self.thumbnail_status_label.setObjectName("statusMeta")
        self.thumbnail_status_label.setVisible(False)

        self.file_path_label.setObjectName("statusMeta")
        self.row_count_label.setObjectName("statusMeta")
        self.modified_label.setObjectName("statusMeta")
        self.drop_result_label.setObjectName("statusMeta")

        status_bar = QStatusBar(self)
        status_bar.addPermanentWidget(self.file_path_label, 2)
        status_bar.addPermanentWidget(self.row_count_label)
        status_bar.addPermanentWidget(self.modified_label)
        status_bar.addPermanentWidget(self.drop_progress_bar)
        status_bar.addPermanentWidget(self.drop_result_label, 2)
        status_bar.addPermanentWidget(self.thumbnail_status_label)
        status_bar.addPermanentWidget(self.thumbnail_progress_bar)
        self.setStatusBar(status_bar)

    def _build_menu_bar(self) -> None:
        self.menuBar().clear()
        self.menuBar().setVisible(True)

    def _build_menus(self) -> None:
        self.file_menu.clear()
        self.file_menu.addAction(self.new_action)
        self.file_menu.addAction(self.open_action)
        self.file_menu.addMenu(self.recent_files_menu)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.save_action)
        self.file_menu.addAction(self.save_as_action)

        self.edit_menu.clear()
        self.edit_menu.addAction(self.undo_action)
        self.edit_menu.addAction(self.redo_action)
        self.edit_menu.addSeparator()
        self.edit_menu.addAction(self.copy_action)
        self.edit_menu.addAction(self.paste_action)
        self.edit_menu.addSeparator()
        self.edit_menu.addAction(self.add_row_action)
        self.edit_menu.addAction(self.delete_row_action)
        self.edit_menu.addSeparator()
        self.edit_menu.addAction(self.regenerate_thumbnails_action)
        self.edit_menu.addAction(self.preferences_action)

        self.sort_menu.clear()
        self.sort_menu.addAction(self.restore_default_sort_action)

        self._build_view_menu()

        self.help_menu.clear()
        self.help_menu.addAction(self.check_updates_action)
        self.help_menu.addSeparator()
        self.help_menu.addAction(self.license_info_action)

        self.menuBar().clear()
        self.menuBar().addMenu(self.file_menu)
        self.menuBar().addMenu(self.edit_menu)
        self.menuBar().addMenu(self.sort_menu)
        self.menuBar().addMenu(self.view_menu)
        self.menuBar().addMenu(self.help_menu)
        self._refresh_recent_files_menu()

    def _build_view_menu(self) -> None:
        self.view_menu.clear()
        self.column_visibility_actions.clear()
        hidden_columns = self._load_hidden_columns()
        for column, header in enumerate(CSV_HEADERS):
            action = QAction(header, self)
            action.setCheckable(True)
            # 初期チェック状態の設定で toggled が発火して設定を上書きしないようブロックする。
            was_blocked = action.blockSignals(True)
            action.setChecked(column not in hidden_columns)
            action.blockSignals(was_blocked)
            action.toggled.connect(
                lambda checked, col=column: self._on_column_visibility_toggled(col, checked)
            )
            self.column_visibility_actions[column] = action
            self.view_menu.addAction(action)
        self.view_menu.addSeparator()
        show_all_action = QAction("すべての列を表示", self)
        show_all_action.triggered.connect(self._show_all_columns)
        self.view_menu.addAction(show_all_action)

    def _connect_signals(self) -> None:
        self.table_view.clearRequested.connect(self.clear_selected_cells)
        self.table_view.addRowRequested.connect(self.add_row)
        self.table_view.deleteRowsRequested.connect(self.delete_selected_rows)
        self.table_view.copyRequested.connect(self.copy_selected_cells)
        self.table_view.pasteRequested.connect(self.paste_cells_from_clipboard)
        self.table_view.pathsDropped.connect(self.handle_dropped_paths)
        self.table_view.dragStateChanged.connect(self._set_drag_feedback)
        self.table_view.customContextMenuRequested.connect(self._open_table_context_menu)
        self.table_view.doubleClicked.connect(self._on_cell_double_clicked)
        self.table_view.horizontalHeader().sectionClicked.connect(self._toggle_sort_by_column)
        self.table_view.horizontalHeader().filterButtonClicked.connect(self._open_column_popup)
        self.table_view.horizontalHeader().sectionResized.connect(self._on_column_resized)
        self.table_view.verticalHeader().sectionResized.connect(self._on_row_resized)

        self.model.modifiedChanged.connect(self._update_all_status)
        self.model.actualRowCountChanged.connect(self._update_all_status)
        self.model.contentChanged.connect(lambda *_: self._update_all_status())
        self.model.modelReset.connect(self._update_all_status)
        self.model.rowsInserted.connect(lambda *_: self._update_all_status())
        self.model.rowsRemoved.connect(lambda *_: self._update_all_status())
        self.model.layoutChanged.connect(self._update_all_status)
        self.proxy_model.modelReset.connect(self._update_all_status)
        self.proxy_model.rowsInserted.connect(lambda *_: self._update_all_status())
        self.proxy_model.rowsRemoved.connect(lambda *_: self._update_all_status())
        self.history.canUndoChanged.connect(self.undo_action.setEnabled)
        self.history.canRedoChanged.connect(self.redo_action.setEnabled)

    def _connect_theme_signals(self) -> None:
        app = QApplication.instance()
        if app is None:
            return
        try:
            app.styleHints().colorSchemeChanged.connect(lambda *_: self._schedule_theme_style_refresh())
        except AttributeError:
            pass
        self.history.cleanChanged.connect(lambda clean: self.model.set_modified(not clean))

    def _update_all_status(self) -> None:
        self._update_window_title()
        self._update_status_labels()
        self._update_summary_memo()
        self.table_view.horizontalHeader().set_filtered_columns(self.proxy_model.filtered_columns())

    def _update_window_title(self) -> None:
        suffix = ""
        if self.current_file_path:
            suffix = f" - {self.current_file_path}"
        if self.model.is_modified():
            suffix = f"{suffix} *"
        self.setWindowTitle(f"{WINDOW_TITLE}{suffix}")

    def _update_status_labels(self) -> None:
        current_path = self.current_file_path or "未作成"
        self.file_path_label.setText(f"ファイル: {current_path}")
        total_rows = self.model.actual_row_count()
        visible_rows = self.proxy_model.rowCount()
        if total_rows == 0:
            visible_rows = 0
        if self.proxy_model.has_active_filters():
            self.row_count_label.setText(f"行数: {total_rows} (表示 {visible_rows})")
        else:
            self.row_count_label.setText(f"行数: {total_rows}")

        self.modified_label.setText("状態: 未保存" if self.model.is_modified() else "状態: 保存済み")
        self.drop_result_label.setText(f"D&D: {self.last_drop_summary}")
        self._update_summary_ribbon()

    def _build_summary_ribbon(self) -> QWidget:
        ribbon = QWidget(self)
        ribbon.setObjectName("summaryRibbon")
        outer_layout = QVBoxLayout(ribbon)
        outer_layout.setContentsMargins(10, 8, 10, 8)
        outer_layout.setSpacing(6)
        outer_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        header_row = RibbonToggleRow(ribbon)
        header_row.setObjectName("summaryToggleRow")
        header_row.setFixedHeight(18)
        header_row.clicked.connect(self._toggle_summary_ribbon)
        header_layout = QHBoxLayout(header_row)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)

        self.summary_toggle_button = RibbonToggleButton(ribbon)
        self.summary_toggle_button.setObjectName("summaryToggleButton")
        self.summary_toggle_button.setChecked(True)
        self.summary_toggle_button.setToolTip("リボンを閉じる")
        self.summary_toggle_button.clicked.connect(lambda *_: self._toggle_summary_ribbon())
        header_layout.addWidget(self.summary_toggle_button)
        header_layout.addStretch(1)
        outer_layout.addWidget(header_row, 0, Qt.AlignmentFlag.AlignTop)

        clip = QWidget(ribbon)
        clip.setObjectName("summaryRibbonClip")
        self.summary_ribbon_clip = clip

        body = QWidget(clip)
        body.setObjectName("summaryRibbonBody")
        self.summary_ribbon_body = body
        layout = QGridLayout(body)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setHorizontalSpacing(8)
        layout.setVerticalSpacing(6)

        metric_positions = (
            ("total_cuts", "総カット数", 0, 0),
            ("delivered", "納品済み", 0, 1),
            ("remaining_delivery", "残り納品", 0, 2),
            ("total_tp", "総TP数", 1, 0),
            ("tp_done", "TP入れ", 1, 1),
            ("remaining_tp", "残りTP数", 1, 2),
            ("total_bg", "総BG数", 2, 0),
            ("bg_done", "BG入れ", 2, 1),
            ("remaining_bg", "残りBG数", 2, 2),
            ("shared", "兼用カット", 0, 3),
            ("bank", "BANK数", 1, 3),
            ("missing", "欠番数", 2, 3),
        )
        for key, title, row, column in metric_positions:
            label = QLabel(ribbon)
            label.setObjectName("summaryMetric")
            label.setMinimumWidth(118)
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            label.setProperty("title", title)
            if key in {"shared", "bank", "missing"}:
                label.setProperty("statusKind", key)
            self.summary_labels[key] = label
            layout.addWidget(label, row, column)

        memo_edit = QPlainTextEdit(body)
        memo_edit.setObjectName("summaryMemo")
        memo_edit.setPlaceholderText("メモ")
        memo_edit.setMinimumWidth(220)
        memo_edit.setMaximumHeight(88)
        memo_edit.textChanged.connect(self._commit_summary_memo)
        self.summary_memo_edit = memo_edit
        layout.addWidget(memo_edit, 0, 4, 3, 1)

        layout.setColumnStretch(4, 1)
        outer_layout.addWidget(clip)
        QTimer.singleShot(0, self._initialize_summary_ribbon_body_geometry)
        return ribbon

    def _toggle_summary_ribbon(self) -> None:
        if self.summary_ribbon_clip is None or self.summary_ribbon_body is None or self.summary_toggle_button is None:
            return
        visible = not self.summary_toggle_button.isChecked()
        self.summary_toggle_button.setChecked(visible)
        self._set_summary_ribbon_visible(visible)
        self.summary_toggle_button.update()
        self.summary_toggle_button.setToolTip("リボンを閉じる" if visible else "リボンを開く")

    def _set_summary_ribbon_visible(self, visible: bool) -> None:
        if self.summary_ribbon_clip is None or self.summary_ribbon_body is None:
            return
        clip = self.summary_ribbon_clip
        body = self.summary_ribbon_body
        expanded_height = body.sizeHint().height()
        if visible:
            clip.setVisible(True)
            body.setVisible(True)
            self._set_summary_ribbon_body_geometry(expanded_height)
            clip.setMinimumHeight(expanded_height)
            clip.setMaximumHeight(expanded_height)
            body.move(0, 0)
        else:
            clip.setMinimumHeight(0)
            clip.setMaximumHeight(0)
            body.move(0, -expanded_height)
            clip.setVisible(False)

    def _initialize_summary_ribbon_body_geometry(self) -> None:
        if self.summary_ribbon_clip is None or self.summary_ribbon_body is None:
            return
        height = self.summary_ribbon_body.sizeHint().height()
        self.summary_ribbon_clip.setMinimumHeight(height)
        self.summary_ribbon_clip.setMaximumHeight(height)
        self._set_summary_ribbon_body_geometry(height)

    def _set_summary_ribbon_body_geometry(self, height: int) -> None:
        if self.summary_ribbon_clip is None or self.summary_ribbon_body is None:
            return
        width = max(self.summary_ribbon_clip.width(), self.summary_ribbon_body.sizeHint().width())
        self.summary_ribbon_body.setGeometry(0, self.summary_ribbon_body.y(), width, height)

    def _update_summary_ribbon(self) -> None:
        if not self.summary_labels:
            return

        stats = self._calculate_cut_summary()
        for key, value in stats.items():
            label = self.summary_labels.get(key)
            if label is None:
                continue
            title = str(label.property("title") or "")
            label.setText(f"{title}: {value}")

    def _calculate_cut_summary(self) -> dict[str, int]:
        return calculate_cut_summary(self.model.rows())

    def _update_summary_memo(self) -> None:
        if self.summary_memo_edit is None:
            return

        memo = self._load_episode_memo()

        self._updating_summary_memo = True
        try:
            if self.summary_memo_edit.toPlainText() != memo:
                self.summary_memo_edit.setPlainText(memo)
            self.summary_memo_edit.setEnabled(True)
        finally:
            self._updating_summary_memo = False

    def _commit_summary_memo(self) -> None:
        if self._updating_summary_memo or self.summary_memo_edit is None:
            return
        self.settings.setValue(self._episode_memo_key(), self.summary_memo_edit.toPlainText())
        self.settings.sync()

    def _load_episode_memo(self) -> str:
        value = self.settings.value(self._episode_memo_key(), "")
        return "" if value is None else str(value)

    def _episode_memo_key(self) -> str:
        if self.current_file_path:
            path_key = self._normalize_recent_path(self.current_file_path).replace("\\", "/")
        else:
            path_key = "__unsaved__"
        return f"{self.EPISODE_MEMO_PREFIX}{path_key}"

    def _open_column_popup(self, column: int) -> None:
        values = self.model.unique_column_values(column)
        allowed_values = self.proxy_model.allowed_values(column)
        checked_values = set(values) if allowed_values is None else allowed_values

        popup = ColumnFilterPopup(
            CSV_HEADERS[column],
            values,
            checked_values,
            self,
        )

        header = self.table_view.horizontalHeader()
        popup_position = header.viewport().mapToGlobal(
            QPoint(header.sectionViewportPosition(column), header.height())
        )
        popup.move(popup_position)

        if popup.exec() != QDialog.DialogCode.Accepted:
            return

        self._apply_column_filter(column, popup.selected_values(), popup.all_values())
        self._update_all_status()

    def _toggle_sort_by_column(self, column: int) -> None:
        if self._sort_column == column:
            order = (
                Qt.SortOrder.DescendingOrder
                if self._sort_order == Qt.SortOrder.AscendingOrder
                else Qt.SortOrder.AscendingOrder
            )
        else:
            order = Qt.SortOrder.AscendingOrder
        self._apply_sort(column, order)

    def _apply_sort(
        self,
        column: int,
        order: Qt.SortOrder,
        *,
        mark_modified: bool = True,
    ) -> None:
        if not 0 <= column < len(CSV_HEADERS):
            return
        self._sort_column = column
        self._sort_order = order
        self.model.sort(column, order)
        if mark_modified:
            self.model.set_modified(True)
        self._update_sort_indicator()

    def _restore_default_sort(self, *, mark_modified: bool = True) -> None:
        self._apply_sort(
            COLUMN_CUT_NUMBER,
            Qt.SortOrder.AscendingOrder,
            mark_modified=mark_modified,
        )

    def _apply_column_filter(
        self,
        column: int,
        selected_values: set[str],
        all_values: set[str],
    ) -> None:
        if selected_values == all_values:
            self.proxy_model.clear_allowed_values(column)
        else:
            self.proxy_model.set_allowed_values(column, selected_values)
        self._update_status_labels()

    def _update_sort_indicator(self) -> None:
        header = self.table_view.horizontalHeader()
        header.setSortIndicatorShown(True)
        header.setSortIndicator(self._sort_column, self._sort_order)
        self.restore_default_sort_action.setEnabled(not self._is_default_sort())

    def _is_default_sort(self) -> bool:
        return (
            self._sort_column == COLUMN_CUT_NUMBER
            and self._sort_order == Qt.SortOrder.AscendingOrder
        )

    def _restore_header_state(self) -> None:
        header = self.table_view.horizontalHeader()
        self._restoring_header_state = True
        try:
            stored_state = self.settings.value(self.HEADER_STATE_KEY)
            if stored_state is not None:
                if header.restoreState(stored_state):
                    return
            self._apply_default_header_order()
        finally:
            self._restoring_header_state = False

    def _apply_default_header_order(self) -> None:
        header = self.table_view.horizontalHeader()
        default_order = [
            "カット番号",
            "AB分け",
            "区分",
            "メモ",
            "TP入れ回数",
            "TP入れ日",
            "BG入れ回数",
            "BG入れ日",
            "テイク",
            "テイク番号",
            "納品日",
            "Roll",
            "動画パス",
            "サムネイル",
        ]
        for target_visual_index, header_name in enumerate(default_order):
            try:
                logical_index = CSV_HEADERS.index(header_name)
            except ValueError:
                continue
            current_visual_index = header.visualIndex(logical_index)
            if current_visual_index != target_visual_index:
                header.moveSection(current_visual_index, target_visual_index)
        self._save_header_state()

    def _save_header_state(self, *_args) -> None:
        if self._restoring_header_state:
            return
        self.settings.setValue(self.HEADER_STATE_KEY, self.table_view.horizontalHeader().saveState())
        self.settings.sync()

    def _on_column_resized(self, logical_index: int, _old_size: int, new_size: int) -> None:
        # リサイズした列が選択範囲に含まれていれば、選択中の全列を同じ幅に揃える。
        if self._syncing_section_size:
            return
        selected_columns = self._selected_columns()
        if logical_index not in selected_columns or len(selected_columns) <= 1:
            return
        header = self.table_view.horizontalHeader()
        self._syncing_section_size = True
        try:
            for column in selected_columns:
                if column != logical_index:
                    header.resizeSection(column, new_size)
        finally:
            self._syncing_section_size = False
        self._save_header_state()

    def _on_row_resized(self, logical_index: int, _old_size: int, new_size: int) -> None:
        # リサイズした行が選択範囲に含まれていれば、選択中の全行を同じ高さに揃える。
        if self._syncing_section_size:
            return
        selected_rows = self._selected_rows()
        if logical_index not in selected_rows or len(selected_rows) <= 1:
            return
        header = self.table_view.verticalHeader()
        self._syncing_section_size = True
        try:
            for row in selected_rows:
                if row != logical_index:
                    header.resizeSection(row, new_size)
        finally:
            self._syncing_section_size = False

    def _selected_columns(self) -> set[int]:
        return {
            index.column()
            for index in self.table_view.selectionModel().selectedIndexes()
            if index.isValid()
        }

    def _selected_rows(self) -> set[int]:
        return {
            index.row()
            for index in self.table_view.selectionModel().selectedIndexes()
            if index.isValid()
        }

    def _on_column_visibility_toggled(self, column: int, visible: bool) -> None:
        self.table_view.setColumnHidden(column, not visible)
        self._save_hidden_columns()

    def _show_all_columns(self) -> None:
        for column, action in self.column_visibility_actions.items():
            if not action.isChecked():
                action.setChecked(True)  # toggled シグナルで列表示と保存が走る
            else:
                self.table_view.setColumnHidden(column, False)
        self._save_hidden_columns()

    def _apply_column_visibility(self) -> None:
        hidden_columns = self._load_hidden_columns()
        for column in range(len(CSV_HEADERS)):
            self.table_view.setColumnHidden(column, column in hidden_columns)
            action = self.column_visibility_actions.get(column)
            if action is not None:
                was_blocked = action.blockSignals(True)
                action.setChecked(column not in hidden_columns)
                action.blockSignals(was_blocked)

    def _load_hidden_columns(self) -> set[int]:
        stored_value = self.settings.value(self.COLUMN_VISIBILITY_KEY, [])
        if stored_value is None:
            return set()
        if isinstance(stored_value, str):
            candidates = [stored_value]
        else:
            candidates = list(stored_value)
        hidden: set[int] = set()
        for candidate in candidates:
            try:
                column = int(candidate)
            except (TypeError, ValueError):
                continue
            if 0 <= column < len(CSV_HEADERS):
                hidden.add(column)
        return hidden

    def _save_hidden_columns(self) -> None:
        hidden = [
            str(column)
            for column in range(len(CSV_HEADERS))
            if self.table_view.isColumnHidden(column)
        ]
        self.settings.setValue(self.COLUMN_VISIBILITY_KEY, hidden)
        self.settings.sync()

    def _schedule_resort(self, changed_columns: list[int]) -> None:
        if self._sort_column not in changed_columns or self._pending_resort:
            return
        self._pending_resort = True
        QTimer.singleShot(0, self._apply_pending_resort)

    def _apply_pending_resort(self) -> None:
        self._pending_resort = False
        self._apply_sort(self._sort_column, self._sort_order, mark_modified=False)

    def create_new_csv(self) -> None:
        if not self._confirm_discard_or_save():
            return

        target_path = self._choose_save_path()
        if not target_path:
            return

        try:
            save_csv_file(target_path, [])
        except CsvLoadError as exc:
            self._show_error("新規 CSV の作成に失敗しました。", str(exc))
            return

        self.model.replace_rows([], modified=False)
        self.history.clear()
        self._set_current_file_path(target_path)
        self._push_recent_file(target_path)
        self.last_drop_summary = "-"
        self._reset_view_state(preserve_row_order=True)
        self.table_view.setFocus()
        self.statusBar().showMessage("新規 CSV を作成しました。", 4000)
        self._update_all_status()

    def open_csv_dialog(self) -> None:
        start_dir = str(Path(self.current_file_path).parent) if self.current_file_path else str(Path.cwd())
        file_path, _ = QFileDialog.getOpenFileName(self, "CSV を開く", start_dir, CSV_FILE_FILTER)
        if not file_path:
            return
        self.open_csv_path(file_path)

    def open_csv_path(self, file_path: str) -> bool:
        return self._load_csv_path(file_path, confirm_unsaved=True, interactive=True)

    def _load_csv_path(
        self,
        file_path: str,
        *,
        confirm_unsaved: bool,
        interactive: bool,
    ) -> bool:
        if confirm_unsaved and not self._confirm_discard_or_save():
            return False

        try:
            load_result = load_csv_file(file_path)
        except CsvLoadError as exc:
            if interactive:
                self._show_error("CSV の読み込みに失敗しました。", str(exc))
            else:
                self.statusBar().showMessage("前回の CSV を復元できませんでした。", 5000)
            return False

        self.model.replace_rows(
            load_result.rows,
            modified=False,
            sort_column=COLUMN_CUT_NUMBER,
            sort_order=Qt.SortOrder.AscendingOrder,
        )
        self.history.clear()
        self._set_current_file_path(file_path)
        self._push_recent_file(file_path)
        self.last_drop_summary = "-"
        self._reset_view_state(preserve_row_order=True)
        self.table_view.setFocus()
        self._update_all_status()

        if interactive:
            self.statusBar().showMessage("CSV を読み込みました。", 4000)
            if load_result.warnings:
                QMessageBox.warning(self, "ヘッダー警告", "\n".join(load_result.warnings))
        else:
            self.statusBar().showMessage("前回の CSV を復元しました。", 4000)

        return True

    def save_csv(self) -> bool:
        if not self.current_file_path:
            return self.save_csv_as()

        if self._pending_resort:
            self._apply_pending_resort()

        try:
            save_csv_file(self.current_file_path, self.model.rows())
        except CsvLoadError as exc:
            self._show_error("CSV の保存に失敗しました。", str(exc))
            return False

        self.history.set_clean()
        self._push_recent_file(self.current_file_path)
        self.statusBar().showMessage("CSV を保存しました。", 4000)
        self._update_all_status()
        return True

    def save_csv_as(self) -> bool:
        target_path = self._choose_save_path(self.current_file_path)
        if not target_path:
            return False

        previous_path = self.current_file_path
        previous_episode_memo = self._load_episode_memo()
        self._set_current_file_path(target_path)
        self.settings.setValue(self._episode_memo_key(), previous_episode_memo)
        if not self.save_csv():
            self._set_current_file_path(previous_path)
            self._update_all_status()
            return False
        return True

    def add_row(self, insert_at: int | None = None) -> None:
        if isinstance(insert_at, bool):
            insert_at = None
        if insert_at is None:
            current_index = self.table_view.currentIndex()
            if current_index.isValid():
                source_index = self.proxy_model.mapToSource(current_index)
                insert_at = source_index.row() if source_index.isValid() else self.model.actual_row_count()
            else:
                insert_at = self.model.actual_row_count()

        source_index = self.model.insert_blank_row(insert_at)
        target_index = self.proxy_model.mapFromSource(source_index)

        if not target_index.isValid():
            self.proxy_model.clear_all_filters()
            self.statusBar().showMessage("行追加のため絞り込みを解除しました。", 3000)
            target_index = self.proxy_model.mapFromSource(source_index)

        if target_index.isValid():
            self.table_view.setCurrentIndex(target_index)
            self.table_view.scrollTo(target_index)
            self.table_view.edit(target_index)

    def add_row_above(self) -> None:
        self.add_row(self._context_row_insert_position(offset=0))

    def add_row_below(self) -> None:
        self.add_row(self._context_row_insert_position(offset=1))

    def delete_selected_rows(self) -> None:
        source_rows = self._selected_source_rows()
        if not source_rows:
            return

        removed_count = self.model.remove_rows_by_numbers(source_rows)
        if removed_count:
            self.statusBar().showMessage(f"{removed_count} 行を削除しました。", 4000)
            self.table_view.setFocus()

    def clear_selected_cells(self) -> None:
        selected_indexes = self.table_view.selectionModel().selectedIndexes()
        if not selected_indexes:
            return

        cleared_count = self.model.clear_indexes(
            [
                self.proxy_model.mapToSource(index)
                for index in selected_indexes
                if index.isValid()
            ]
        )
        if cleared_count:
            self.statusBar().showMessage(f"{cleared_count} セルをクリアしました。", 3000)

    def _on_cell_double_clicked(self, proxy_index) -> None:
        if not proxy_index.isValid():
            return
        source_index = self.proxy_model.mapToSource(proxy_index)
        if not source_index.isValid():
            return
        column = source_index.column()
        if column == COLUMN_VIDEO_PATH:
            self._reveal_video_path(self.model.video_path_for_row(source_index.row()))
        elif column == COLUMN_THUMBNAIL:
            self._play_video(self.model.video_path_for_row(source_index.row()))

    def regenerate_thumbnails(self) -> None:
        paths: list[str] = []
        seen: set[str] = set()
        for row in range(self.model.actual_row_count()):
            path_text = self.model.video_path_for_row(row)
            if not path_text:
                continue
            key = path_text.casefold()
            if key in seen:
                continue
            seen.add(key)
            paths.append(path_text)

        if not paths:
            self.statusBar().showMessage("再生成できる動画パスがありません。", 4000)
            return

        # 既存キャッシュ（メモリ状態）を破棄し、全件を強制再生成する（ffmpeg が上書き）。
        self.thumbnail_provider.invalidate(paths)
        for path_text in paths:
            self.thumbnail_provider.request(path_text, force=True)
        self.table_view.viewport().update()
        self.statusBar().showMessage(f"サムネイルを再生成しています（{len(paths)} 件）。", 4000)

    def _on_thumbnail_progress(self, done: int, total: int) -> None:
        # 生成中はスケルトンのライトスイープを動かし、フリーズしていないことを示す。
        self.cut_delegate.set_skeleton_animating(total > 0)
        if total <= 0:
            self.thumbnail_progress_bar.setVisible(False)
            self.thumbnail_status_label.setVisible(False)
            return
        self.thumbnail_progress_bar.setRange(0, total)
        self.thumbnail_progress_bar.setValue(done)
        self.thumbnail_progress_bar.setVisible(True)
        self.thumbnail_status_label.setText(f"サムネイル生成 {done}/{total}")
        self.thumbnail_status_label.setVisible(True)

    def _reveal_video_path(self, video_path: str) -> None:
        path_text = str(video_path or "").strip()
        if not path_text or not Path(path_text).exists():
            self.statusBar().showMessage("動画ファイルが見つかりません。", 4000)
            return
        native_path = str(Path(path_text))
        if sys.platform.startswith("win"):
            # ファイルを選択した状態でエクスプローラーを開く。
            subprocess.Popen(["explorer", f"/select,{native_path}"])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", native_path])
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(native_path).parent)))

    def _play_video(self, video_path: str) -> None:
        path_text = str(video_path or "").strip()
        if not path_text or not Path(path_text).exists():
            self.statusBar().showMessage("動画ファイルが見つかりません。", 4000)
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path_text)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path_text])
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(path_text))
        except OSError as exc:
            self._show_error("動画を再生できませんでした。", str(exc))

    def undo(self) -> None:
        if self.history.undo():
            self._update_sort_indicator()
            self.table_view.setFocus()

    def redo(self) -> None:
        if self.history.redo():
            self._update_sort_indicator()
            self.table_view.setFocus()

    def copy_selected_cells(self) -> None:
        selected_indexes = [index for index in self.table_view.selectionModel().selectedIndexes() if index.isValid()]
        if not selected_indexes:
            return

        min_row = min(index.row() for index in selected_indexes)
        max_row = max(index.row() for index in selected_indexes)
        min_column = min(index.column() for index in selected_indexes)
        max_column = max(index.column() for index in selected_indexes)

        selected_values = {
            (index.row(), index.column()): str(self.proxy_model.data(index, Qt.ItemDataRole.EditRole) or "")
            for index in selected_indexes
        }
        lines: list[str] = []
        for row in range(min_row, max_row + 1):
            line = [
                selected_values.get((row, column), "")
                for column in range(min_column, max_column + 1)
            ]
            lines.append("\t".join(line))

        QApplication.clipboard().setText("\n".join(lines))
        self.statusBar().showMessage(f"{len(selected_indexes)} セルをコピーしました。", 3000)

    def paste_cells_from_clipboard(self) -> None:
        clipboard_text = QApplication.clipboard().text()
        matrix = self._clipboard_matrix(clipboard_text)
        if not matrix:
            return

        selected_indexes = [index for index in self.table_view.selectionModel().selectedIndexes() if index.isValid()]
        if len(matrix) == 1 and len(matrix[0]) == 1 and len(selected_indexes) > 1:
            pasted_cells = 0
            for proxy_index in selected_indexes:
                source_index = self.proxy_model.mapToSource(proxy_index)
                if not source_index.isValid():
                    continue
                if self.model.setData(source_index, matrix[0][0], Qt.ItemDataRole.EditRole):
                    pasted_cells += 1
            if pasted_cells:
                self.statusBar().showMessage(f"{pasted_cells} セルに貼り付けました。", 3000)
            return

        current_index = self.table_view.currentIndex()
        if current_index.isValid():
            source_start_index = self.proxy_model.mapToSource(current_index)
        else:
            source_start_index = self.model.index(0, 0)

        start_row = source_start_index.row() if source_start_index.isValid() else 0
        start_column = source_start_index.column() if source_start_index.isValid() else 0
        new_rows = self.model.rows()
        required_row_count = start_row + len(matrix)
        while len(new_rows) < required_row_count:
            new_rows.append([""] * len(CSV_HEADERS))

        pasted_cells = 0
        for row_offset, row_values in enumerate(matrix):
            target_row = start_row + row_offset
            for column_offset, value in enumerate(row_values):
                target_column = start_column + column_offset
                if target_column >= len(CSV_HEADERS):
                    break
                if target_column in NON_DATA_COLUMNS:
                    continue
                text = "" if value is None else str(value)
                if new_rows[target_row][target_column] == text:
                    continue
                new_rows[target_row][target_column] = text
                pasted_cells += 1

        if pasted_cells:
            self.model.replace_rows(
                new_rows,
                modified=True,
                sort_column=self._sort_column,
                sort_order=self._sort_order,
            )
            target_proxy_index = self.proxy_model.mapFromSource(self.model.index(start_row, start_column))
            if target_proxy_index.isValid():
                self.table_view.setCurrentIndex(target_proxy_index)
                self.table_view.scrollTo(target_proxy_index)
            self.statusBar().showMessage(f"{pasted_cells} セルを貼り付けました。", 3000)

    def _open_table_context_menu(self, position: QPoint) -> None:
        clicked_index = self.table_view.indexAt(position)
        if clicked_index.isValid():
            self.table_view.setCurrentIndex(clicked_index)

        has_selection = bool(self.table_view.selectionModel().selectedIndexes())
        self.copy_action.setEnabled(has_selection)
        self.paste_action.setEnabled(bool(self._clipboard_matrix(QApplication.clipboard().text())))
        has_reference_row = self._context_row_insert_position(offset=0) is not None
        self.add_row_above_action.setEnabled(has_reference_row)
        self.add_row_below_action.setEnabled(True)
        self.clear_values_action.setEnabled(has_selection)
        self.delete_row_action.setEnabled(bool(self._selected_source_rows()))

        menu = QMenu(self)
        menu.addAction(self.copy_action)
        menu.addAction(self.paste_action)
        menu.addSeparator()
        menu.addAction(self.clear_values_action)
        menu.addAction(self.delete_row_action)
        menu.addSeparator()
        menu.addAction(self.add_row_above_action)
        menu.addAction(self.add_row_below_action)
        menu.exec(self.table_view.viewport().mapToGlobal(position))

    def _context_row_insert_position(self, *, offset: int) -> int | None:
        current_index = self.table_view.currentIndex()
        if current_index.isValid():
            source_index = self.proxy_model.mapToSource(current_index)
            if source_index.isValid():
                return source_index.row() + offset

        if self.model.actual_row_count() == 0:
            return 0 if offset == 1 else None

        return self.model.actual_row_count()

    @staticmethod
    def _clipboard_matrix(text: str) -> list[list[str]]:
        normalized = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
        if not normalized:
            return []

        lines = normalized.split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        return [line.split("\t") for line in lines] if lines else []

    def open_settings_dialog(self) -> None:
        dialog = SettingsDialog(
            self.history.limit,
            self.shortcut_manager.all_sequences(),
            self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        new_limit = dialog.undo_limit()
        self.history.set_limit(new_limit)
        self._save_undo_limit(new_limit)
        self.shortcut_manager.update(dialog.shortcuts())
        self._apply_shortcuts()
        self.statusBar().showMessage(f"アンドゥ履歴数を {new_limit} 回に設定しました。", 4000)

    def check_for_updates(self) -> None:
        if self._update_check_thread is not None:
            self.statusBar().showMessage("すでに更新確認を実行中です。", 3000)
            return

        self.check_updates_action.setEnabled(False)
        self.statusBar().showMessage("更新を確認しています...", 0)

        self._update_check_thread = QThread(self)
        self._update_check_worker = UpdateCheckWorker()
        self._update_check_worker.moveToThread(self._update_check_thread)
        self._update_check_thread.started.connect(self._update_check_worker.run)
        self._update_check_worker.finished.connect(self._on_update_check_finished)
        self._update_check_worker.failed.connect(self._on_update_check_failed)
        self._update_check_worker.finished.connect(self._update_check_thread.quit)
        self._update_check_worker.failed.connect(self._update_check_thread.quit)
        self._update_check_thread.finished.connect(self._cleanup_update_check)
        self._update_check_thread.start()

    def _on_update_check_finished(self, result: UpdateCheckResult) -> None:
        self.statusBar().showMessage("更新確認が完了しました。", 4000)
        self._show_update_check_result(result)

    def _on_update_check_failed(self, message: str) -> None:
        self.statusBar().showMessage("更新確認に失敗しました。", 5000)
        QMessageBox.warning(self, "更新確認", message)

    def _cleanup_update_check(self) -> None:
        if self._update_check_worker is not None:
            self._update_check_worker.deleteLater()
        if self._update_check_thread is not None:
            self._update_check_thread.deleteLater()
        self._update_check_worker = None
        self._update_check_thread = None
        self.check_updates_action.setEnabled(True)

    def _show_update_check_result(self, result: UpdateCheckResult) -> None:
        release = result.release
        asset = release.asset

        message_box = QMessageBox(self)
        message_box.setWindowTitle("更新確認")
        message_box.setDetailedText(release.body or "リリースノートはありません。")

        info_lines = [
            f"現在: {result.current_version}",
            f"最新: {release.version}",
        ]
        if release.published_at and release.published_at != "-":
            info_lines.append(f"公開日: {release.published_at}")
        if asset is not None:
            info_lines.append(f"配布ファイル: {asset.name} ({human_readable_size(asset.size)})")
        else:
            info_lines.append("配布ファイル: 自動更新に使える asset が見つかりませんでした")

        if result.update_available:
            message_box.setIcon(QMessageBox.Icon.Information)
            message_box.setText(f"新しいバージョン {release.version} が見つかりました。")
            message_box.setInformativeText("\n".join(info_lines))
            update_button = None
            if asset is not None:
                update_label = "ダウンロードして更新"
                if asset.suffix == ".exe":
                    update_label = "インストーラーを起動"
                update_button = message_box.addButton(update_label, QMessageBox.ButtonRole.AcceptRole)
            open_release_button = message_box.addButton("リリースページを開く", QMessageBox.ButtonRole.ActionRole)
            close_button = message_box.addButton("閉じる", QMessageBox.ButtonRole.RejectRole)
            message_box.setDefaultButton(close_button)
            message_box.exec()

            clicked = message_box.clickedButton()
            if clicked == update_button and asset is not None:
                self._download_update_asset(asset)
            elif clicked == open_release_button:
                self._open_release_page(release.html_url)
            return

        message_box.setIcon(QMessageBox.Icon.Information)
        message_box.setText(f"現在のバージョン {result.current_version} は最新です。")
        message_box.setInformativeText("\n".join(info_lines))
        open_release_button = message_box.addButton("リリースページを開く", QMessageBox.ButtonRole.ActionRole)
        close_button = message_box.addButton("閉じる", QMessageBox.ButtonRole.AcceptRole)
        message_box.setDefaultButton(close_button)
        message_box.exec()
        if message_box.clickedButton() == open_release_button:
            self._open_release_page(release.html_url)

    def _download_update_asset(self, asset: UpdateAsset) -> None:
        if self._update_download_thread is not None:
            self.statusBar().showMessage("更新ファイルをすでにダウンロード中です。", 3000)
            return

        self._download_progress_dialog = QProgressDialog("更新ファイルをダウンロードしています...", "", 0, 0, self)
        self._download_progress_dialog.setWindowTitle("更新をダウンロード")
        self._download_progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self._download_progress_dialog.setAutoClose(False)
        self._download_progress_dialog.setAutoReset(False)
        self._download_progress_dialog.setMinimumDuration(0)
        self._download_progress_dialog.setCancelButton(None)
        self._download_progress_dialog.show()

        self._update_download_thread = QThread(self)
        self._update_download_worker = UpdateDownloadWorker(asset)
        self._update_download_worker.moveToThread(self._update_download_thread)
        self._update_download_thread.started.connect(self._update_download_worker.run)
        self._update_download_worker.progress.connect(self._on_update_download_progress)
        self._update_download_worker.finished.connect(self._on_update_download_finished)
        self._update_download_worker.failed.connect(self._on_update_download_failed)
        self._update_download_worker.finished.connect(self._update_download_thread.quit)
        self._update_download_worker.failed.connect(self._update_download_thread.quit)
        self._update_download_thread.finished.connect(self._cleanup_update_download)
        self._update_download_thread.start()

    def _on_update_download_progress(self, downloaded_bytes: int, total_bytes: int) -> None:
        if self._download_progress_dialog is None:
            return

        if total_bytes > 0:
            self._download_progress_dialog.setMaximum(total_bytes)
            self._download_progress_dialog.setValue(min(downloaded_bytes, total_bytes))
            self._download_progress_dialog.setLabelText(
                "更新ファイルをダウンロードしています...\n"
                f"{human_readable_size(downloaded_bytes)} / {human_readable_size(total_bytes)}"
            )
        else:
            self._download_progress_dialog.setMaximum(0)
            self._download_progress_dialog.setLabelText(
                "更新ファイルをダウンロードしています...\n"
                f"{human_readable_size(downloaded_bytes)}"
            )

    def _on_update_download_finished(self, downloaded_path: str) -> None:
        if self._download_progress_dialog is not None:
            self._download_progress_dialog.close()

        path = Path(downloaded_path)
        try:
            prepared_update = prepare_update(path)
        except UpdateError as exc:
            self.statusBar().showMessage("更新ファイルの準備に失敗しました。", 5000)
            QMessageBox.warning(
                self,
                "更新準備",
                f"{exc}\n\nダウンロード先:\n{path}",
            )
            self._open_local_path(path.parent)
            return

        if not self._confirm_ready_for_restart(prepared_update):
            self.statusBar().showMessage("更新は保留しました。", 4000)
            self._open_local_path(path.parent)
            return

        try:
            self._launch_prepared_update(prepared_update)
        except UpdateError as exc:
            self._skip_close_confirmation = False
            self._show_error("更新の起動に失敗しました。", str(exc))
            self._open_local_path(path.parent)

    def _on_update_download_failed(self, message: str) -> None:
        if self._download_progress_dialog is not None:
            self._download_progress_dialog.close()
        self.statusBar().showMessage("更新ファイルのダウンロードに失敗しました。", 5000)
        QMessageBox.warning(self, "更新ダウンロード", message)

    def _cleanup_update_download(self) -> None:
        if self._update_download_worker is not None:
            self._update_download_worker.deleteLater()
        if self._update_download_thread is not None:
            self._update_download_thread.deleteLater()
        self._update_download_worker = None
        self._update_download_thread = None
        if self._download_progress_dialog is not None:
            self._download_progress_dialog.deleteLater()
        self._download_progress_dialog = None

    def _confirm_ready_for_restart(self, prepared_update: PreparedUpdate) -> bool:
        if not self._confirm_discard_or_save():
            return False

        mode_label = "インストーラーを起動" if prepared_update.mode == "installer" else "更新を適用"
        answer = QMessageBox.question(
            self,
            "更新を適用",
            f"{mode_label}するため、CutManager を終了します。続行しますか。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        return answer == QMessageBox.StandardButton.Yes

    def _launch_prepared_update(self, prepared_update: PreparedUpdate) -> None:
        self._skip_close_confirmation = True
        launch_result = QProcess.startDetached(prepared_update.launch_program, prepared_update.launch_arguments)
        launched = launch_result[0] if isinstance(launch_result, tuple) else bool(launch_result)
        if not launched:
            raise UpdateError("更新プロセスを起動できませんでした。")

        self.statusBar().showMessage("更新を開始します。アプリを終了します。", 4000)
        QApplication.instance().quit()

    def show_license_info(self) -> None:
        notices = self._load_bundled_text("THIRD_PARTY_NOTICES.txt", "THIRD_PARTY_NOTICES.md")
        ffmpeg_license_path = self._find_bundled_resource("ffmpeg-LICENSE.txt")

        message_box = QMessageBox(self)
        message_box.setWindowTitle("ライセンス情報")
        message_box.setIcon(QMessageBox.Icon.Information)
        message_box.setText(
            "CutManager は MIT License で提供されています。\n"
            "動画サムネイル生成に FFmpeg (GPLv3) を外部プログラムとして使用しています。"
        )
        message_box.setInformativeText(
            "FFmpeg プロジェクト: https://ffmpeg.org/\n"
            "対応ソース入手先: https://ffmpeg.org/download.html"
        )
        if notices:
            message_box.setDetailedText(notices)

        if ffmpeg_license_path is not None:
            open_ffmpeg_license = message_box.addButton(
                "FFmpeg ライセンス文を開く", QMessageBox.ButtonRole.ActionRole
            )
        else:
            open_ffmpeg_license = None
        open_source_button = message_box.addButton("FFmpeg ソース入手先を開く", QMessageBox.ButtonRole.ActionRole)
        close_button = message_box.addButton("閉じる", QMessageBox.ButtonRole.AcceptRole)
        message_box.setDefaultButton(close_button)
        message_box.exec()

        clicked = message_box.clickedButton()
        if open_ffmpeg_license is not None and clicked == open_ffmpeg_license:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(ffmpeg_license_path)))
        elif clicked == open_source_button:
            QDesktopServices.openUrl(QUrl("https://ffmpeg.org/download.html"))

    def _load_bundled_text(self, *names: str) -> str:
        for name in names:
            resource = self._find_bundled_resource(name)
            if resource is not None:
                try:
                    return resource.read_text(encoding="utf-8")
                except OSError:
                    continue
        return ""

    @staticmethod
    def _find_bundled_resource(name: str) -> Path | None:
        # 同梱データファイル（onefile 展開先 / 実行ファイル同ディレクトリ / リポジトリ）を探す。
        search_dirs: list[Path] = []

        def _add(path: Path | None) -> None:
            if path is not None and path not in search_dirs:
                search_dirs.append(path)

        try:
            _add(Path(__file__).resolve().parent.parent)
        except Exception:
            pass
        try:
            _add(Path(sys.executable).resolve().parent)
        except Exception:
            pass
        if sys.argv and sys.argv[0]:
            try:
                _add(Path(sys.argv[0]).resolve().parent)
            except Exception:
                pass
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            _add(Path(meipass))

        for directory in search_dirs:
            candidate = directory / name
            if candidate.is_file():
                return candidate
        return None

    @staticmethod
    def _open_release_page(url: str | None = None) -> None:
        target_url = QUrl(str(url or RELEASES_PAGE_URL))
        if target_url.isValid():
            QDesktopServices.openUrl(target_url)

    @staticmethod
    def _open_local_path(path: Path) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        # 受理可否の判定（ファイルシステム stat を含む）はドラッグ開始時に一度だけ行い、
        # 移動中は結果を使い回す（ホバー中の重さを防ぐ）。
        self._drag_accept_cache = self._can_accept_paths(self._extract_drop_paths(event))
        if self._drag_accept_cache:
            self._set_drag_feedback(True)
            event.acceptProposedAction()
            return
        self._set_drag_feedback(False)
        event.ignore()

    def dragMoveEvent(self, event) -> None:
        accepted = self._drag_accept_cache
        if accepted is None:
            accepted = self._can_accept_paths(self._extract_drop_paths(event))
            self._drag_accept_cache = accepted
        if accepted:
            self._set_drag_feedback(True)
            event.acceptProposedAction()
            return
        self._set_drag_feedback(False)
        event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self._drag_accept_cache = None
        self._set_drag_feedback(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self._drag_accept_cache = None
        self._set_drag_feedback(False)
        if self.handle_dropped_paths(self._extract_drop_paths(event)):
            event.acceptProposedAction()
            return
        event.ignore()

    def handle_dropped_paths(self, paths: list[Path]) -> bool:
        normalized_paths = self._normalize_drop_paths(paths)
        if not normalized_paths:
            return False

        drop_type = self._classify_drop_paths(normalized_paths)
        if drop_type == "unsupported":
            QMessageBox.information(
                self,
                "ドロップ不可",
                "CSV/.cutmgr ファイル 1 件、素材フォルダー、PSD/PSB ファイル、"
                "動画ファイル、または動画を含む Roll フォルダーをドロップしてください。",
            )
            return False

        progress_message = {
            "csv": "CSV を開いています...",
            "folders": "素材を取り込んでいます...",
            "videos": "動画情報を反映しています...",
        }[drop_type]

        self._show_drop_progress(progress_message)
        try:
            if drop_type == "csv":
                return self.open_csv_path(str(normalized_paths[0]))

            if drop_type == "folders":
                return self.import_material_folders(normalized_paths)

            return self.import_video_files(normalized_paths)
        finally:
            self._hide_drop_progress()

    def import_material_folders(self, folders: list[Path]) -> bool:
        if not self.current_file_path:
            answer = QMessageBox.question(
                self,
                "保存先が未設定です",
                "素材を取り込む前に、新規 CSV の保存先を指定しますか。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return False

            self.create_new_csv()
            if not self.current_file_path:
                return False

        import_date = QDate.currentDate().toString(IMPORT_DATE_FORMAT)

        try:
            result = build_rows_from_dropped_folders(folders, self.model.cut_keys(), import_date)
        except OSError as exc:
            self._show_error("フォルダーの読み込みに失敗しました。", str(exc))
            return False
        except ValueError as exc:
            self._show_error("フォルダーの読み込みに失敗しました。", str(exc))
            return False

        if result.rows or result.updates:
            merged_rows = apply_material_updates(self.model.rows(), result.updates)
            merged_rows.extend(result.rows)
            self.model.replace_rows(
                merged_rows,
                modified=True,
                sort_column=self._sort_column,
                sort_order=self._sort_order,
            )
            self._update_sort_indicator()

        self.last_drop_summary = f"素材追加 {result.added_count} / 既存更新 {result.updated_count} / 抽出失敗 {result.failed_count}"
        self.statusBar().showMessage(self.last_drop_summary, 7000)
        self._update_status_labels()
        return True

    def import_video_files(self, video_paths: list[Path]) -> bool:
        if not self.current_file_path:
            QMessageBox.information(self, "動画反映", "先に CSV を開くか新規作成してください。")
            return False

        # ドロップされたフォルダー（Roll フォルダー等）は中の動画ファイルへ展開する。
        video_paths = self._gather_video_files(video_paths)
        if not video_paths:
            QMessageBox.information(
                self,
                "動画反映",
                "ドロップしたフォルダー／ファイルに動画が見つかりませんでした。",
            )
            return False

        delivery_date = QDate.currentDate().toString(IMPORT_DATE_FORMAT)
        created_from_videos = False

        if self.model.actual_row_count() == 0:
            draft_result = build_rows_from_video_files(video_paths, self.model.cut_keys(), delivery_date)
            if draft_result.added_count == 0:
                QMessageBox.information(
                    self,
                    "動画反映",
                    "動画名からカット番号を読み取れなかったため登録できませんでした。",
                )
                self.last_drop_summary = f"動画仮登録 0 / 読み取り失敗 {draft_result.failed_count}"
                self.statusBar().showMessage(self.last_drop_summary, 7000)
                self._update_status_labels()
                return False

            answer = QMessageBox.question(
                self,
                "動画反映",
                (
                    "この CSV はまだ空です。\n"
                    f"動画 {len(video_paths)} 件から {draft_result.added_count} カットを仮登録して、"
                    "納品情報を反映しますか。"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Yes,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return False

            self.model.replace_rows(
                draft_result.rows,
                modified=True,
                sort_column=self._sort_column,
                sort_order=self._sort_order,
            )
            self._update_sort_indicator()
            created_from_videos = True

        result = apply_videos_to_rows(video_paths, self.model.rows(), delivery_date)

        if result.updated_count:
            self.model.replace_rows(
                result.rows,
                modified=True,
                sort_column=self._sort_column,
                sort_order=self._sort_order,
            )
            self._update_sort_indicator()

        if created_from_videos:
            self.last_drop_summary = f"動画仮登録 {result.updated_count} / 読み取り失敗 {result.failed_count}"
        else:
            self.last_drop_summary = (
                f"動画反映 {result.updated_count} / 未一致 {result.unmatched_count} / 抽出失敗 {result.failed_count}"
            )
            if result.unmatched_files:
                unmatched_lines = "\n".join(result.unmatched_files)
                message_box = QMessageBox(self)
                message_box.setIcon(QMessageBox.Icon.Information)
                message_box.setWindowTitle("動画反映")
                message_box.setText(
                    (
                        f"{result.updated_count} 件反映しました。\n"
                        f"{len(result.unmatched_files)} 件は一致するカットが CSV にないため反映されませんでした。"
                    )
                )
                message_box.setInformativeText(f"未一致ファイル:\n{unmatched_lines}")
                register_button = message_box.addButton("未一致を仮登録", QMessageBox.ButtonRole.ActionRole)
                close_button = message_box.addButton("閉じる", QMessageBox.ButtonRole.AcceptRole)
                message_box.setDefaultButton(close_button)
                message_box.exec()

                if message_box.clickedButton() == register_button:
                    draft_result = build_rows_from_video_files(video_paths, self.model.cut_keys(), delivery_date)
                    if draft_result.added_count:
                        merged_rows = self.model.rows()
                        merged_rows.extend(draft_result.rows)
                        self.model.replace_rows(
                            merged_rows,
                            modified=True,
                            sort_column=self._sort_column,
                            sort_order=self._sort_order,
                        )
                        self._update_sort_indicator()
                    self.last_drop_summary = (
                        f"動画反映 {result.updated_count} / 未一致 {result.unmatched_count} / "
                        f"仮登録 {draft_result.added_count} / 抽出失敗 {result.failed_count}"
                    )
        self.statusBar().showMessage(self.last_drop_summary, 7000)
        self._update_status_labels()
        return True

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._skip_close_confirmation:
            event.accept()
            return
        if self._confirm_discard_or_save():
            event.accept()
            return
        event.ignore()

    def open_recent_file(self, file_path: str | None = None) -> None:
        if not file_path:
            return

        normalized_path = str(Path(file_path))
        if not Path(normalized_path).exists():
            QMessageBox.information(self, "最近開いたファイル", f"ファイルが見つかりません。\n{normalized_path}")
            self._remove_recent_file(normalized_path)
            return

        self.open_csv_path(normalized_path)

    def _confirm_discard_or_save(self) -> bool:
        if not self.model.is_modified():
            return True

        message_box = QMessageBox(self)
        message_box.setWindowTitle("未保存の変更")
        message_box.setText("未保存の変更があります。保存しますか。")
        save_button = message_box.addButton("保存", QMessageBox.ButtonRole.AcceptRole)
        discard_button = message_box.addButton("破棄", QMessageBox.ButtonRole.DestructiveRole)
        cancel_button = message_box.addButton("キャンセル", QMessageBox.ButtonRole.RejectRole)
        message_box.setDefaultButton(save_button)
        message_box.exec()

        clicked = message_box.clickedButton()
        if clicked == save_button:
            return self.save_csv()
        if clicked == discard_button:
            return True
        if clicked == cancel_button:
            return False
        return False

    def _selected_source_rows(self) -> list[int]:
        indexes = self.table_view.selectionModel().selectedIndexes()
        rows = []
        for index in indexes:
            source_index = self.proxy_model.mapToSource(index)
            if source_index.isValid() and source_index.row() < self.model.actual_row_count():
                rows.append(source_index.row())
        return sorted(set(rows))

    def _choose_save_path(self, suggested_path: str | None = None) -> str | None:
        if suggested_path:
            start_path = suggested_path
        elif self.current_file_path:
            start_path = self.current_file_path
        else:
            start_path = str(Path.cwd() / f"cut_list{PROJECT_FILE_EXTENSION}")

        # 既定選択フィルターを .cutmgr 専用にし、拡張子未指定時の既定を .cutmgr に固定する。
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存先を選択",
            start_path,
            CSV_FILE_FILTER,
            PROJECT_SAVE_FILTER,
        )
        if not file_path:
            return None

        normalized_path = self._normalize_csv_path(file_path)
        path_obj = Path(normalized_path)
        if path_obj.exists():
            answer = QMessageBox.question(
                self,
                "上書き確認",
                f"{normalized_path}\nを上書きしますか。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return None

        return normalized_path

    def _show_drop_progress(self, message: str) -> None:
        self.drop_progress_bar.setVisible(True)
        self.statusBar().showMessage(message)
        QApplication.processEvents()

    def _hide_drop_progress(self) -> None:
        self.drop_progress_bar.setVisible(False)
        QApplication.processEvents()

    def _set_drag_feedback(self, active: bool) -> None:
        # 状態が変わらないドラッグ移動中は、重いスタイル再適用を行わない。
        if self._drag_feedback_active == active:
            return
        self._drag_feedback_active = active
        self.drop_hint_label.setVisible(active)
        self._apply_theme_styles()

    def changeEvent(self, event) -> None:
        if event.type() in (QEvent.Type.PaletteChange, QEvent.Type.ApplicationPaletteChange):
            self._schedule_theme_style_refresh()
        super().changeEvent(event)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._update_drop_hint_geometry()

    def _schedule_theme_style_refresh(self) -> None:
        if self._theme_apply_pending:
            return
        self._theme_apply_pending = True
        QTimer.singleShot(0, self._apply_theme_styles)

    def _update_drop_hint_geometry(self) -> None:
        viewport = self.table_view.viewport()
        if viewport is None:
            return
        margin = 16
        hint_height = 54
        width = max(220, viewport.width() - (margin * 2))
        self.drop_hint_label.setGeometry(margin, margin, width, hint_height)
        self.drop_hint_label.raise_()

    def _apply_theme_styles(self) -> None:
        if self._applying_theme_styles:
            return
        self._theme_apply_pending = False
        self._applying_theme_styles = True
        palette = self.palette()
        try:
            if self._is_dark_theme():
                # Reuse the docs dark theme colors.
                base = QColor("#0f172a")
                alternate = QColor("#162033")
                text = QColor("#e5eefc")
                muted = QColor("#c7d2e5")
                highlight = QColor("#3b82f6")
                highlighted_text = QColor("#eff6ff")
                mid = QColor("#334155")
                button = QColor("#172033")
                button_text = QColor("#bfdbfe")
                paper = QColor("#172033")
                surface = QColor("#10192b")
            else:
                base = QColor("#ffffff")
                alternate = QColor("#f7faff")
                text = QColor("#0f172a")
                muted = QColor("#475569")
                highlight = QColor("#2563eb")
                highlighted_text = QColor("#eff6ff")
                mid = QColor("#cbd5e1")
                button = QColor("#ffffff")
                button_text = QColor("#0f172a")
                paper = QColor("#ffffff")
                surface = QColor("#f8fbff")

            border_color = highlight if self._drag_feedback_active else mid
            border_width = 2 if self._drag_feedback_active else 1
            table_background = self._blend_colors(base, highlight, 0.08) if self._drag_feedback_active else base
            selection_background = self._blend_colors(base, highlight, 0.22)
            hint_background = self._blend_colors(button, highlight, 0.18)
            hint_text = highlight if self._is_color_dark(hint_background) == self._is_color_dark(highlight) else button_text

            self.drop_hint_label.setStyleSheet(
                "QLabel {"
                "padding: 10px 14px;"
                f"border: 1px dashed {border_color.name()};"
                f"background: {hint_background.name()};"
                f"color: {hint_text.name()};"
                "border-radius: 8px;"
                "font-weight: 600;"
                "}"
            )

            window_stylesheet = (
                "QMainWindow {"
                f"background: {surface.name()};"
                f"color: {text.name()};"
                "}"
                "QWidget#mainContainer {"
                f"background: {surface.name()};"
                "}"
                "QWidget#summaryRibbon {"
                f"background: {paper.name()};"
                f"border-bottom: 1px solid {self._blend_colors(mid, base, 0.35).name()};"
                "}"
                "QWidget#summaryRibbonBody {"
                f"background: {paper.name()};"
                "}"
                "QWidget#summaryRibbonClip {"
                f"background: {paper.name()};"
                "}"
                "QLabel#summaryMetric {"
                f"background: {self._blend_colors(base, highlight, 0.08).name()};"
                f"color: {text.name()};"
                f"border: 1px solid {self._blend_colors(mid, base, 0.20).name()};"
                "border-radius: 6px;"
                "padding: 6px 8px;"
                "font-weight: 600;"
                "}"
                "QLabel#summaryMetric[statusKind=\"shared\"] {"
                f"background: {self._blend_colors(base, QColor('#22c55e'), 0.22 if self._is_dark_theme() else 0.18).name()};"
                f"color: {text.name()};"
                "}"
                "QLabel#summaryMetric[statusKind=\"bank\"] {"
                f"background: {self._blend_colors(base, QColor('#ef4444'), 0.30 if self._is_dark_theme() else 0.18).name()};"
                f"color: {text.name()};"
                "}"
                "QLabel#summaryMetric[statusKind=\"missing\"] {"
                f"background: {self._blend_colors(base, QColor('#1e3a8a' if self._is_dark_theme() else '#64748b'), 0.50 if self._is_dark_theme() else 0.28).name()};"
                f"color: {(highlighted_text if self._is_dark_theme() else text).name()};"
                "}"
                "QPlainTextEdit#summaryMemo {"
                f"background: {paper.name()};"
                f"color: {text.name()};"
                f"border: 1px solid {self._blend_colors(mid, base, 0.20).name()};"
                "border-radius: 6px;"
                "padding: 6px 8px;"
                "font-weight: 500;"
                "}"
                "QWidget#summaryToggleButton {"
                "background: transparent;"
                f"color: {muted.name()};"
                "border: 0px;"
                "border-radius: 0px;"
                "padding: 0px;"
                "font-weight: 600;"
                "}"
                "QWidget#summaryToggleButton:hover {"
                "background: transparent;"
                f"color: {text.name()};"
                "}"
                "QMenuBar {"
                f"background: {paper.name()};"
                f"color: {text.name()};"
                f"border: 1px solid {self._blend_colors(mid, base, 0.35).name()};"
                "border-radius: 8px;"
                "padding: 4px 6px;"
                "spacing: 8px;"
                "}"
                "QMenuBar::item {"
                "padding: 6px 10px;"
                "border-radius: 6px;"
                "background: transparent;"
                "}"
                "QMenuBar::item:selected {"
                f"background: {self._blend_colors(base, highlight, 0.12).name()};"
                f"color: {text.name()};"
                "}"
                "QMenu {"
                f"background: {paper.name()};"
                f"color: {text.name()};"
                f"border: 1px solid {mid.name()};"
                "border-radius: 8px;"
                "padding: 6px;"
                "}"
                "QMenu::item {"
                "padding: 7px 12px;"
                "border-radius: 6px;"
                "margin: 2px 0;"
                "}"
                "QMenu::item:selected {"
                f"background: {self._blend_colors(base, highlight, 0.16).name()};"
                f"color: {text.name()};"
                "}"
                "QStatusBar {"
                f"background: {paper.name()};"
                f"color: {muted.name()};"
                f"border-top: 1px solid {self._blend_colors(mid, base, 0.35).name()};"
                "padding: 4px 8px;"
                "}"
                "QStatusBar::item { border: 0; }"
                "QLabel#statusMeta {"
                f"color: {muted.name()};"
                "padding: 0 4px;"
                "font-weight: 500;"
                "}"
                "QHeaderView::section {"
                f"background: {paper.name()};"
                f"color: {text.name()};"
                f"border: 0px;"
                f"border-bottom: 1px solid {mid.name()};"
                "padding: 10px 12px;"
                "font-weight: 600;"
                "}"
                "QTableView QHeaderView::section:vertical {"
                f"background: {paper.name()};"
                f"color: {muted.name()};"
                f"border: 0px;"
                f"border-right: 1px solid {mid.name()};"
                f"border-bottom: 1px solid {mid.name()};"
                "padding: 2px 6px;"
                "font-weight: 500;"
                "}"
                "QTableCornerButton::section {"
                f"background: {paper.name()};"
                f"border: 0px;"
                f"border-right: 1px solid {mid.name()};"
                f"border-bottom: 1px solid {mid.name()};"
                "}"
                "QProgressBar {"
                f"background: {self._blend_colors(base, mid, 0.10).name()};"
                f"border: 1px solid {mid.name()};"
                "border-radius: 6px;"
                "padding: 1px;"
                "}"
                "QProgressBar::chunk {"
                f"background: {highlight.name()};"
                "border-radius: 4px;"
                "}"
                "QLineEdit, QComboBox, QSpinBox, QListWidget, QPlainTextEdit {"
                f"background: {paper.name()};"
                f"color: {text.name()};"
                f"border: 1px solid {mid.name()};"
                "border-radius: 6px;"
                "padding: 6px 10px;"
                "selection-background-color: " + highlight.name() + ";"
                "selection-color: " + highlighted_text.name() + ";"
                "}"
                "QComboBox::drop-down {"
                "border: 0px;"
                "width: 24px;"
                "}"
                "QPushButton {"
                f"background: {paper.name()};"
                f"color: {text.name()};"
                f"border: 1px solid {mid.name()};"
                "border-radius: 6px;"
                "padding: 7px 12px;"
                "font-weight: 600;"
                "}"
                "QPushButton:hover {"
                f"background: {self._blend_colors(base, highlight, 0.10).name()};"
                "}"
                "QPushButton:pressed {"
                f"background: {self._blend_colors(base, highlight, 0.18).name()};"
                "}"
                "QDialog {"
                f"background: {surface.name()};"
                f"color: {text.name()};"
                "}"
            )
            window_styles_changed = window_stylesheet != self._last_window_stylesheet
            if window_stylesheet != self._last_window_stylesheet:
                self.setStyleSheet(window_stylesheet)
                self._last_window_stylesheet = window_stylesheet

            table_stylesheet = (
                "QTableView {"
                f"border: {border_width}px solid {border_color.name()};"
                f"background: {table_background.name()};"
                f"alternate-background-color: {alternate.name()};"
                f"color: {text.name()};"
                f"gridline-color: {mid.name()};"
                f"selection-background-color: {selection_background.name()};"
                f"selection-color: {text.name()};"
                "border-radius: 0px;"
                "padding: 0px;"
                "}"
                "QTableView::item {"
                "padding: 1px 3px;"
                "border: 0px;"
                "margin: 0px;"
                "}"
                "QTableView::item:selected {"
                f"background: {selection_background.name()};"
                f"color: {text.name()};"
                "border: 0px;"
                "outline: none;"
                "}"
                "QTableView::item:selected:active {"
                f"background: {selection_background.name()};"
                f"color: {text.name()};"
                "border: 0px;"
                "outline: none;"
                "}"
                "QTableView::item:focus { outline: none; }"
            )
            table_styles_changed = table_stylesheet != self._last_table_stylesheet
            if table_stylesheet != self._last_table_stylesheet:
                self.table_view.setStyleSheet(table_stylesheet)
                self._last_table_stylesheet = table_stylesheet

            if window_styles_changed or table_styles_changed:
                self.model.refresh_colors()
            if window_styles_changed:
                self.menuBar().update()
                self.statusBar().update()
            if table_styles_changed:
                self.table_view.horizontalHeader().viewport().update()
                self.table_view.verticalHeader().viewport().update()
        finally:
            self._applying_theme_styles = False

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

    def _is_dark_theme(self) -> bool:
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
        return self._is_color_dark(self.palette().color(QPalette.ColorRole.Base))

    def _reset_view_state(self, *, preserve_row_order: bool = False) -> None:
        self._pending_resort = False
        self.proxy_model.clear_all_filters()
        self._sort_column = COLUMN_CUT_NUMBER
        self._sort_order = Qt.SortOrder.AscendingOrder
        if not preserve_row_order:
            self.model.sort(self._sort_column, self._sort_order)
        self._update_sort_indicator()

    def _restore_last_session_file(self) -> None:
        last_file_path = self._load_last_session_file()
        if not last_file_path:
            return

        normalized_path = self._normalize_recent_path(last_file_path)
        if not Path(normalized_path).exists():
            self._clear_last_session_file()
            self._remove_recent_file(normalized_path)
            return

        if not self._load_csv_path(normalized_path, confirm_unsaved=False, interactive=False):
            self._clear_last_session_file()

    def _load_last_session_file(self) -> str | None:
        stored_value = self.settings.value(self.LAST_SESSION_FILE_KEY)
        if not stored_value:
            return None
        return str(stored_value)

    def _set_current_file_path(self, file_path: str | None) -> None:
        self.current_file_path = file_path
        if file_path:
            normalized_path = self._normalize_recent_path(file_path)
            self.settings.setValue(self.LAST_SESSION_FILE_KEY, normalized_path)
        else:
            self._clear_last_session_file()
        self.settings.sync()

    def _clear_last_session_file(self) -> None:
        self.settings.remove(self.LAST_SESSION_FILE_KEY)
        self.settings.sync()

    def _load_recent_files(self) -> list[str]:
        stored_value = self.settings.value("recentFiles", [])
        if stored_value is None:
            return []
        if isinstance(stored_value, str):
            candidates = [stored_value]
        else:
            candidates = list(stored_value)
        return self._normalize_recent_files(candidates)

    def _load_undo_limit(self) -> int:
        stored_value = self.settings.value(self.UNDO_LIMIT_KEY, self.DEFAULT_UNDO_LIMIT)
        try:
            return max(10, int(stored_value))
        except (TypeError, ValueError):
            return self.DEFAULT_UNDO_LIMIT

    def _save_undo_limit(self, undo_limit: int) -> None:
        self.settings.setValue(self.UNDO_LIMIT_KEY, int(undo_limit))
        self.settings.sync()

    def _save_recent_files(self) -> None:
        self.settings.setValue("recentFiles", self.recent_files)
        self.settings.sync()

    def _push_recent_file(self, file_path: str | None) -> None:
        if not file_path:
            return
        updated = [self._normalize_recent_path(file_path), *self.recent_files]
        self.recent_files = self._normalize_recent_files(updated)
        self._save_recent_files()
        self._refresh_recent_files_menu()

    def _remove_recent_file(self, file_path: str) -> None:
        normalized_path = self._normalize_recent_path(file_path)
        self.recent_files = [
            path for path in self.recent_files if path.casefold() != normalized_path.casefold()
        ]
        self._save_recent_files()
        self._refresh_recent_files_menu()

    def _refresh_recent_files_menu(self) -> None:
        self.recent_files_menu.clear()

        if not self.recent_files:
            placeholder_action = self.recent_files_menu.addAction("最近開いたファイルはありません")
            placeholder_action.setEnabled(False)
            return

        for path in self.recent_files:
            label = self._format_recent_file_label(path)
            action = self.recent_files_menu.addAction(label)
            action.setToolTip(path)
            action.triggered.connect(lambda checked=False, p=path: self.open_recent_file(p))

    @classmethod
    def _normalize_recent_files(cls, paths: list[str]) -> list[str]:
        normalized: list[str] = []
        seen: set[str] = set()

        for file_path in paths:
            if not file_path:
                continue
            normalized_path = cls._normalize_recent_path(file_path)
            key = normalized_path.casefold()
            if key in seen:
                continue
            seen.add(key)
            normalized.append(normalized_path)
            if len(normalized) >= cls.MAX_RECENT_FILES:
                break

        return normalized

    @staticmethod
    def _normalize_recent_path(file_path: str) -> str:
        return QDir.toNativeSeparators(str(Path(file_path).resolve(strict=False)))

    @staticmethod
    def _format_recent_file_label(file_path: str) -> str:
        path = Path(file_path)
        parent_text = str(path.parent)
        return f"{path.name} | {parent_text}"

    @staticmethod
    def _normalize_csv_path(file_path: str) -> str:
        normalized = QDir.toNativeSeparators(file_path)
        lowered = normalized.casefold()
        if any(lowered.endswith(extension) for extension in SUPPORTED_PROJECT_EXTENSIONS):
            return normalized
        # 拡張子が無い場合は独自拡張子 .cutmgr を既定にする。
        return f"{normalized}{PROJECT_FILE_EXTENSION}"

    @staticmethod
    def _extract_drop_paths(event) -> list[Path]:
        mime_data = event.mimeData()
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
    def _normalize_drop_paths(paths: list[Path]) -> list[Path]:
        normalized: list[Path] = []
        seen: set[str] = set()

        for path in paths:
            key = str(path.resolve(strict=False)).casefold()
            if key in seen:
                continue
            seen.add(key)
            normalized.append(path)

        return normalized

    @staticmethod
    def _classify_drop_paths(paths: list[Path]) -> str:
        if not paths:
            return "unsupported"

        csv_paths = [
            path
            for path in paths
            if path.is_file() and path.suffix.casefold() in SUPPORTED_PROJECT_EXTENSIONS
        ]
        if len(csv_paths) == 1 and len(paths) == 1:
            return "csv"

        # 動画ファイル、または動画を含むフォルダー（Roll フォルダー等）は動画取り込み。
        video_entries = [path for path in paths if MainWindow._is_video_drop_entry(path)]
        if video_entries and len(video_entries) == len(paths):
            return "videos"

        material_paths = [
            path
            for path in paths
            if path.is_dir() or (path.is_file() and path.suffix.casefold() in BG_FILE_EXTENSIONS)
        ]
        if material_paths and len(material_paths) == len(paths):
            return "folders"
        return "unsupported"

    @staticmethod
    def _is_video_drop_entry(path: Path) -> bool:
        if path.is_file():
            return path.suffix.casefold() in VIDEO_FILE_EXTENSIONS
        if path.is_dir():
            return MainWindow._folder_contains_videos(path)
        return False

    @staticmethod
    def _folder_contains_videos(folder: Path) -> bool:
        # 動画が 1 つでも見つかれば True。巨大フォルダーでの走査コストを抑えるため件数上限を設ける。
        scanned = 0
        try:
            for child in folder.rglob("*"):
                scanned += 1
                if scanned > 5000:
                    break
                if child.suffix.casefold() in VIDEO_FILE_EXTENSIONS:
                    return True
        except OSError:
            return False
        return False

    @staticmethod
    def _gather_video_files(paths: list[Path]) -> list[Path]:
        video_files: list[Path] = []
        seen: set[str] = set()
        for path in paths:
            candidates: list[Path] = []
            if path.is_file() and path.suffix.casefold() in VIDEO_FILE_EXTENSIONS:
                candidates = [path]
            elif path.is_dir():
                try:
                    candidates = sorted(
                        (
                            child
                            for child in path.rglob("*")
                            if child.is_file() and child.suffix.casefold() in VIDEO_FILE_EXTENSIONS
                        ),
                        key=lambda child: str(child).casefold(),
                    )
                except OSError:
                    candidates = []
            for candidate in candidates:
                key = str(candidate.resolve(strict=False)).casefold()
                if key in seen:
                    continue
                seen.add(key)
                video_files.append(candidate)
        return video_files

    def _can_accept_paths(self, paths: list[Path]) -> bool:
        return self._classify_drop_paths(paths) != "unsupported"

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)
