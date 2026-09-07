from __future__ import annotations

CSV_HEADERS = [
    "カット番号",
    "メモ",
    "AB分け",
    "区分",
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

COLUMN_CUT_NUMBER = 0
COLUMN_MEMO = 1
COLUMN_AB_GROUP = 2
COLUMN_STATUS = 3
COLUMN_TP_LOAD_COUNT = 4
COLUMN_TP_DATE = 5
COLUMN_BG_LOAD_COUNT = 6
COLUMN_BG_DATE = 7
COLUMN_TAKE = 8
COLUMN_TAKE_NUMBER = 9
COLUMN_DELIVERY_DATE = 10
COLUMN_ROLL = 11
COLUMN_VIDEO_PATH = 12
COLUMN_THUMBNAIL = 13

COLUMN_MATERIAL_LOAD_COUNT = COLUMN_TP_LOAD_COUNT
COLUMN_MATERIAL_DATE = COLUMN_TP_DATE

STATUS_OPTIONS = ("", "兼用", "BANK", "欠番")
TP_LOAD_COUNT_OPTIONS = ("", "BGOnly", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")
BG_LOAD_COUNT_OPTIONS = ("", "全セル", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")
LEGACY_STATUS_HEADERS = ("兼用", "BANK", "欠番")
STATUS_ROW_BACKGROUND_HEX = {
    "欠番": "#4b5563",
    "兼用": "#dcfce7",
    "BANK": "#fee2e2",
}
STATUS_ROW_FOREGROUND_HEX = {
    "欠番": "#f8fafc",
}

# 独自拡張子 .cutmgr（中身はCSV）。保存時の既定拡張子として使い、
# 開く/D&D では .cutmgr と .csv の両方を受け付ける。
PROJECT_FILE_EXTENSION = ".cutmgr"
SUPPORTED_PROJECT_EXTENSIONS = (".cutmgr", ".csv")
CSV_FILE_FILTER = (
    "CutManager ファイル (*.cutmgr *.csv);;"
    "CutManager プロジェクト (*.cutmgr);;"
    "CSV ファイル (*.csv)"
)
# 保存ダイアログで既定選択にするフィルター。既定拡張子を .cutmgr に固定する。
PROJECT_SAVE_FILTER = "CutManager プロジェクト (*.cutmgr)"
WINDOW_TITLE = "CutManager"
WINDOW_SIZE = (1220, 720)
IMPORT_DATE_FORMAT = "yyyy/MM/dd"

VIDEO_FILE_EXTENSIONS = {
    ".mp4",
    ".mov",
    ".mxf",
    ".avi",
    ".wmv",
    ".m4v",
}

BG_FILE_EXTENSIONS = {
    ".psd",
    ".psb",
}
