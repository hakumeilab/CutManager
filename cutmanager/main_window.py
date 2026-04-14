from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QDate, QDir, QEvent, QPoint, QProcess, QSettings, QThread, QTimer, Qt, QUrl
from PySide6.QtGui import QAction, QCloseEvent, QColor, QDesktopServices, QDragEnterEvent, QDropEvent, QKeySequence, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QDialog,
    QFileDialog,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QProgressDialog,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from .constants import (
    COLUMN_CUT_NUMBER,
    CSV_FILE_FILTER,
    CSV_HEADERS,
    IMPORT_DATE_FORMAT,
    VIDEO_FILE_EXTENSIONS,
    WINDOW_SIZE,
    WINDOW_TITLE,
)
from .csv_io import CsvLoadError, load_csv_file, save_csv_file
from .filter_popup import ColumnFilterPopup
from .folder_import import apply_material_updates, build_rows_from_dropped_folders
from .history import HistoryManager
from .model import CutTableModel
from .proxy import CutFilterProxyModel
from .settings_dialog import SettingsDialog
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


class MainWindow(QMainWindow):
    MAX_RECENT_FILES = 8
    LAST_SESSION_FILE_KEY = "lastSessionFile"
    UNDO_LIMIT_KEY = "undoLimit"
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
        self._theme_apply_pending = False
        self._applying_theme_styles = False
        self._last_window_stylesheet = ""
        self._last_table_stylesheet = ""
        self.settings = QSettings("CutManager", "CutManager")
        self.recent_files = self._load_recent_files()

        self.model = CutTableModel(parent=self)
        self.history = HistoryManager(self._load_undo_limit(), self)
        self.model.set_history_manager(self.history)
        self.proxy_model = CutFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.model)

        self.table_view = CutTableView(self)
        self.drop_hint_label = QLabel(self)
        self.file_path_label = QLabel(self)
        self.row_count_label = QLabel(self)
        self.modified_label = QLabel(self)
        self.drop_result_label = QLabel(self)
        self.drop_progress_bar = QProgressBar(self)
        self.file_menu = QMenu("ファイル", self)
        self.edit_menu = QMenu("編集", self)
        self.sort_menu = QMenu("並べ替え", self)
        self.help_menu = QMenu("ヘルプ", self)
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
        self.preferences_action: QAction
        self.restore_default_sort_action: QAction
        self.check_updates_action: QAction

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
        self.new_action.setShortcut(QKeySequence.StandardKey.New)
        self.new_action.triggered.connect(self.create_new_csv)

        self.open_action = QAction("開く", self)
        self.open_action.setShortcut(QKeySequence.StandardKey.Open)
        self.open_action.triggered.connect(self.open_csv_dialog)

        self.save_action = QAction("上書き保存", self)
        self.save_action.setShortcut(QKeySequence.StandardKey.Save)
        self.save_action.triggered.connect(self.save_csv)

        self.save_as_action = QAction("名前を付けて保存", self)
        self.save_as_action.setShortcut(QKeySequence.StandardKey.SaveAs)
        self.save_as_action.triggered.connect(self.save_csv_as)

        self.undo_action = QAction("元に戻す", self)
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)
        self.undo_action.setEnabled(False)
        self.undo_action.triggered.connect(self.undo)

        self.redo_action = QAction("やり直し", self)
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)
        self.redo_action.setEnabled(False)
        self.redo_action.triggered.connect(self.redo)

        self.copy_action = QAction("コピー", self)
        self.copy_action.setShortcut(QKeySequence.StandardKey.Copy)
        self.copy_action.triggered.connect(self.copy_selected_cells)

        self.paste_action = QAction("貼り付け", self)
        self.paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        self.paste_action.triggered.connect(self.paste_cells_from_clipboard)

        self.add_row_action = QAction("行追加", self)
        self.add_row_action.setShortcut(QKeySequence("Insert"))
        self.add_row_action.triggered.connect(self.add_row)

        self.add_row_above_action = QAction("上に行を追加", self)
        self.add_row_above_action.triggered.connect(self.add_row_above)

        self.add_row_below_action = QAction("下に行を追加", self)
        self.add_row_below_action.triggered.connect(self.add_row_below)

        self.delete_row_action = QAction("行削除", self)
        self.delete_row_action.setShortcut(QKeySequence("Ctrl+Delete"))
        self.delete_row_action.triggered.connect(self.delete_selected_rows)

        self.preferences_action = QAction("環境設定", self)
        self.preferences_action.triggered.connect(self.open_settings_dialog)

        self.restore_default_sort_action = QAction("カット番号順に戻す", self)
        self.restore_default_sort_action.setEnabled(False)
        self.restore_default_sort_action.triggered.connect(self._restore_default_sort)

        self.check_updates_action = QAction("更新を確認", self)
        self.check_updates_action.triggered.connect(self.check_for_updates)

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
            self.preferences_action,
            self.restore_default_sort_action,
            self.check_updates_action,
        ):
            self.addAction(action)

    def _build_ui(self) -> None:
        self.setWindowTitle(WINDOW_TITLE)

        self._build_menu_bar()
        self._build_menus()

        self.drop_hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_hint_label.setVisible(False)
        self.drop_hint_label.setText("ここに CSV / 素材フォルダー / 動画ファイルをドロップ")

        self.table_view.setModel(self.proxy_model)
        self.table_view.setItemDelegate(CutItemDelegate(self.table_view))
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
        header.setSortIndicatorShown(True)
        header.setSortIndicator(self._sort_column, self._sort_order)
        self._apply_theme_styles()

        default_widths = [120, 90, 110, 100, 115, 90, 105, 115]
        for column, width in enumerate(default_widths):
            self.table_view.setColumnWidth(column, width)

        self._set_drag_feedback(False)

        container = QWidget(self)
        container.setObjectName("mainContainer")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.table_view, 1)
        self.setCentralWidget(container)
        self.drop_hint_label.setParent(self.table_view.viewport())
        self.drop_hint_label.raise_()
        self._update_drop_hint_geometry()

        self.drop_progress_bar.setRange(0, 0)
        self.drop_progress_bar.setTextVisible(False)
        self.drop_progress_bar.setVisible(False)
        self.drop_progress_bar.setFixedWidth(160)
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
        self.edit_menu.addAction(self.preferences_action)

        self.sort_menu.clear()
        self.sort_menu.addAction(self.restore_default_sort_action)

        self.help_menu.clear()
        self.help_menu.addAction(self.check_updates_action)

        self.menuBar().clear()
        self.menuBar().addMenu(self.file_menu)
        self.menuBar().addMenu(self.edit_menu)
        self.menuBar().addMenu(self.sort_menu)
        self.menuBar().addMenu(self.help_menu)
        self._refresh_recent_files_menu()

    def _connect_signals(self) -> None:
        self.table_view.clearRequested.connect(self.clear_selected_cells)
        self.table_view.addRowRequested.connect(self.add_row)
        self.table_view.deleteRowsRequested.connect(self.delete_selected_rows)
        self.table_view.copyRequested.connect(self.copy_selected_cells)
        self.table_view.pasteRequested.connect(self.paste_cells_from_clipboard)
        self.table_view.pathsDropped.connect(self.handle_dropped_paths)
        self.table_view.dragStateChanged.connect(self._set_drag_feedback)
        self.table_view.customContextMenuRequested.connect(self._open_table_context_menu)
        self.table_view.horizontalHeader().sectionClicked.connect(self._toggle_sort_by_column)
        self.table_view.horizontalHeader().filterButtonClicked.connect(self._open_column_popup)

        self.model.modifiedChanged.connect(self._update_all_status)
        self.model.actualRowCountChanged.connect(self._update_all_status)
        self.model.contentChanged.connect(self._schedule_resort)
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
        self._set_current_file_path(target_path)
        if not self.save_csv():
            self._set_current_file_path(previous_path)
            self._update_all_status()
            return False
        return True

    def add_row(self, insert_at: int | None = None) -> None:
        if insert_at is None:
            current_index = self.table_view.currentIndex()
            if current_index.isValid():
                source_index = self.proxy_model.mapToSource(current_index)
                insert_at = source_index.row() + 1 if source_index.isValid() else self.model.actual_row_count()
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

        menu = QMenu(self)
        menu.addAction(self.copy_action)
        menu.addAction(self.paste_action)
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
        dialog = SettingsDialog(self.history.limit, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        new_limit = dialog.undo_limit()
        self.history.set_limit(new_limit)
        self._save_undo_limit(new_limit)
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

    @staticmethod
    def _open_release_page(url: str | None = None) -> None:
        target_url = QUrl(str(url or RELEASES_PAGE_URL))
        if target_url.isValid():
            QDesktopServices.openUrl(target_url)

    @staticmethod
    def _open_local_path(path: Path) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if self._can_accept_paths(self._extract_drop_paths(event)):
            self._set_drag_feedback(True)
            event.acceptProposedAction()
            return
        self._set_drag_feedback(False)
        event.ignore()

    def dragMoveEvent(self, event) -> None:
        if self._can_accept_paths(self._extract_drop_paths(event)):
            self._set_drag_feedback(True)
            event.acceptProposedAction()
            return
        self._set_drag_feedback(False)
        event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self._set_drag_feedback(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
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
                "CSV ファイル 1 件、素材フォルダー、または動画ファイルをドロップしてください。",
            )
            return False

        progress_message = {
            "csv": "CSV を開いています...",
            "folders": "素材フォルダーを取り込んでいます...",
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
                "素材フォルダーを取り込む前に、新規 CSV の保存先を指定しますか。",
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

        self.last_drop_summary = (
            f"素材追加 {result.added_count} / 既存更新 {result.updated_count} / 抽出失敗 {result.failed_count}"
        )
        self.statusBar().showMessage(self.last_drop_summary, 7000)
        self._update_status_labels()
        return True

    def import_video_files(self, video_paths: list[Path]) -> bool:
        if not self.current_file_path:
            QMessageBox.information(self, "動画反映", "先に CSV を開くか新規作成してください。")
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
            start_path = str(Path.cwd() / "cut_list.csv")

        file_path, _ = QFileDialog.getSaveFileName(self, "CSV の保存先を選択", start_path, CSV_FILE_FILTER)
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
                "QLineEdit, QComboBox, QSpinBox, QListWidget {"
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
            if table_stylesheet != self._last_table_stylesheet:
                self.table_view.setStyleSheet(table_stylesheet)
                self._last_table_stylesheet = table_stylesheet

            self.model.refresh_colors()
            self.menuBar().update()
            self.statusBar().update()
            self.table_view.horizontalHeader().viewport().update()
            self.table_view.verticalHeader().viewport().update()
            self.table_view.viewport().update()
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
        if normalized.casefold().endswith(".csv"):
            return normalized
        return f"{normalized}.csv"

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

        csv_paths = [path for path in paths if path.is_file() and path.suffix.casefold() == ".csv"]
        folder_paths = [path for path in paths if path.is_dir()]
        video_paths = [path for path in paths if path.is_file() and path.suffix.casefold() in VIDEO_FILE_EXTENSIONS]

        if len(csv_paths) == 1 and len(paths) == 1:
            return "csv"
        if folder_paths and len(folder_paths) == len(paths):
            return "folders"
        if video_paths and len(video_paths) == len(paths):
            return "videos"
        return "unsupported"

    def _can_accept_paths(self, paths: list[Path]) -> bool:
        return self._classify_drop_paths(paths) != "unsupported"

    def _show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)
