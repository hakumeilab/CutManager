from __future__ import annotations

CSV_HEADERS = [
    "カット番号",
    "AB分け",
    "区分",
    "素材入れ回数",
    "素材入れ日",
    "テイク",
    "テイク番号",
    "納品日",
]

COLUMN_CUT_NUMBER = 0
COLUMN_AB_GROUP = 1
COLUMN_STATUS = 2
COLUMN_MATERIAL_LOAD_COUNT = 3
COLUMN_MATERIAL_DATE = 4
COLUMN_TAKE = 5
COLUMN_TAKE_NUMBER = 6
COLUMN_DELIVERY_DATE = 7

STATUS_OPTIONS = ("", "兼用", "BANK", "欠番")
LEGACY_STATUS_HEADERS = ("兼用", "BANK", "欠番")

CSV_FILE_FILTER = "CSV Files (*.csv)"
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
