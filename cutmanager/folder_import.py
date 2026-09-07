from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from .constants import (
    BG_FILE_EXTENSIONS,
    COLUMN_AB_GROUP,
    COLUMN_BG_DATE,
    COLUMN_BG_LOAD_COUNT,
    COLUMN_CUT_NUMBER,
    COLUMN_DELIVERY_DATE,
    COLUMN_STATUS,
    COLUMN_TAKE,
    COLUMN_TAKE_NUMBER,
    COLUMN_TP_DATE,
    COLUMN_TP_LOAD_COUNT,
    CSV_HEADERS,
)


CUT_NUMBER_PATTERN = re.compile(r"(?<!\d)(\d{3})(?!\d)")
CUT_IDENTIFIER_PATTERN = re.compile(r"(?<!\d)(\d{3})([A-Za-z]?)(?![A-Za-z0-9])")
# フォルダー名の "roll01" 等（大文字小文字・区切り記号ゆれを許容）を抽出する。
ROLL_PATTERN = re.compile(r"(?i)roll[ _\-]*(\d+)")
# フォルダー名末尾付近の YYMMDD（6桁）を検出する。
DELIVERY_DATE_PATTERN = re.compile(r"(?<!\d)(\d{2})(\d{2})(\d{2})(?!\d)")


def extract_roll(name: str) -> str:
    """フォルダー/ファイル名から Roll 値（例: "roll01"）を抽出する。無ければ空文字。"""
    match = ROLL_PATTERN.search(str(name or ""))
    if match is None:
        return ""
    return f"roll{match.group(1)}"


def extract_delivery_date(name: str, date_format: str = "%Y/%m/%d") -> str:
    """フォルダー/ファイル名末尾の YYMMDD を YYYY/MM/DD 文字列へ変換する。

    複数該当した場合は最も右（末尾寄り）を採用する。世紀は 00-69→2000 年代、
    70-99→1900 年代に固定する。妥当な日付が無ければ空文字を返す。
    """
    from datetime import date as _date

    last_valid = ""
    for match in DELIVERY_DATE_PATTERN.finditer(str(name or "")):
        year_two, month, day = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
        year = 2000 + year_two if year_two <= 69 else 1900 + year_two
        try:
            resolved = _date(year, month, day)
        except ValueError:
            continue
        last_valid = resolved.strftime(date_format)
    return last_valid


@dataclass(frozen=True, slots=True)
class CutIdentifier:
    cut_number: str
    ab_group: str = ""

    @property
    def key(self) -> tuple[str, str]:
        return make_cut_key(self.cut_number, self.ab_group)


@dataclass(slots=True)
class FolderImportResult:
    rows: list[list[str]]
    added_count: int
    updated_count: int
    failed_count: int
    updates: list["MaterialRowUpdate"]


@dataclass(slots=True)
class MaterialRowUpdate:
    cut_number: str
    ab_group: str
    mark_compatible: bool
    tp_load_increment: int
    tp_date: str
    bg_load_increment: int
    bg_date: str

    @property
    def key(self) -> tuple[str, str]:
        return make_cut_key(self.cut_number, self.ab_group)


def make_cut_key(cut_number: str, ab_group: str = "") -> tuple[str, str]:
    return (str(cut_number or "").strip(), str(ab_group or "").strip().upper())


def extract_cut_number(name: str) -> str | None:
    cut_identifiers = extract_cut_identifiers(name)
    if not cut_identifiers:
        return None
    return cut_identifiers[0].cut_number


def extract_cut_identifiers(name: str) -> list[CutIdentifier]:
    seen: set[CutIdentifier] = set()
    cut_identifiers: list[CutIdentifier] = []

    for match in CUT_IDENTIFIER_PATTERN.finditer(name):
        cut_identifier = CutIdentifier(
            cut_number=match.group(1),
            ab_group=match.group(2).upper(),
        )
        if cut_identifier in seen:
            continue
        seen.add(cut_identifier)
        cut_identifiers.append(cut_identifier)

    return cut_identifiers


def extract_cut_numbers(name: str) -> list[str]:
    return [cut_identifier.cut_number for cut_identifier in extract_cut_identifiers(name)]


def build_rows_from_material_folder(
    folder_path: str | Path,
    existing_cut_keys: set[tuple[str, str]],
    import_date: str,
) -> FolderImportResult:
    return build_rows_from_dropped_folders([folder_path], existing_cut_keys, import_date)


def build_rows_from_dropped_folders(
    folder_paths: list[str | Path],
    existing_cut_keys: set[tuple[str, str]],
    import_date: str,
) -> FolderImportResult:
    seen_existing_cut_keys = {
        make_cut_key(cut_number, ab_group)
        for cut_number, ab_group in existing_cut_keys
        if str(cut_number or "").strip()
    }
    seen_folders: set[str] = set()
    rows_by_cut: dict[tuple[str, str], list[str]] = {}
    updates_by_cut: dict[tuple[str, str], MaterialRowUpdate] = {}
    failed_count = 0

    for folder_path in folder_paths:
        root = Path(folder_path)
        if not root.is_dir() and not root.is_file():
            raise ValueError(f"素材が存在しません: {root}")

        for candidate in _iter_candidate_materials(root):
            folder_key = str(candidate.resolve(strict=False)).casefold()
            if folder_key in seen_folders:
                continue
            seen_folders.add(folder_key)

            cut_identifiers = extract_cut_identifiers(_candidate_name(candidate))
            if not cut_identifiers:
                failed_count += 1
                continue

            is_bg = is_bg_material(candidate)
            is_compatible = len(cut_identifiers) > 1
            for cut_identifier in cut_identifiers:
                cut_key = cut_identifier.key
                if cut_key in seen_existing_cut_keys:
                    update = updates_by_cut.get(cut_key)
                    if update is None:
                        updates_by_cut[cut_key] = MaterialRowUpdate(
                            cut_number=cut_identifier.cut_number,
                            ab_group=cut_identifier.ab_group,
                            mark_compatible=is_compatible,
                            tp_load_increment=0 if is_bg else 1,
                            tp_date="" if is_bg else import_date,
                            bg_load_increment=1 if is_bg else 0,
                            bg_date=import_date if is_bg else "",
                        )
                    else:
                        update.mark_compatible = update.mark_compatible or is_compatible
                        if is_bg:
                            update.bg_load_increment += 1
                            update.bg_date = import_date
                        else:
                            update.tp_load_increment += 1
                            update.tp_date = import_date
                    continue

                row = rows_by_cut.get(cut_key)
                if row is None:
                    rows_by_cut[cut_key] = _build_material_row(cut_identifier, import_date, is_compatible, is_bg)
                    continue

                row[COLUMN_STATUS] = "兼用" if is_compatible else row[COLUMN_STATUS]
                if is_bg:
                    row[COLUMN_BG_LOAD_COUNT] = str(_parse_load_count(row[COLUMN_BG_LOAD_COUNT]) + 1)
                    row[COLUMN_BG_DATE] = import_date
                else:
                    row[COLUMN_TP_LOAD_COUNT] = str(_parse_load_count(row[COLUMN_TP_LOAD_COUNT]) + 1)
                    row[COLUMN_TP_DATE] = import_date

    return FolderImportResult(
        rows=list(rows_by_cut.values()),
        added_count=len(rows_by_cut),
        updated_count=len(updates_by_cut),
        failed_count=failed_count,
        updates=list(updates_by_cut.values()),
    )


def apply_material_updates(rows: list[list[str]], updates: list[MaterialRowUpdate]) -> list[list[str]]:
    if not updates:
        return [_normalize_row(row) for row in rows]

    updated_rows = [_normalize_row(row) for row in rows]
    row_by_cut = {
        make_cut_key(row[COLUMN_CUT_NUMBER], row[COLUMN_AB_GROUP]): index
        for index, row in enumerate(updated_rows)
        if row and row[COLUMN_CUT_NUMBER]
    }

    for update in updates:
        row_index = row_by_cut.get(update.key)
        if row_index is None:
            continue

        row = updated_rows[row_index]
        if update.mark_compatible:
            row[COLUMN_STATUS] = "兼用"
        if update.tp_load_increment:
            row[COLUMN_TP_LOAD_COUNT] = str(_parse_load_count(row[COLUMN_TP_LOAD_COUNT]) + update.tp_load_increment)
            row[COLUMN_TP_DATE] = update.tp_date
        if update.bg_load_increment:
            row[COLUMN_BG_LOAD_COUNT] = str(_parse_load_count(row[COLUMN_BG_LOAD_COUNT]) + update.bg_load_increment)
            row[COLUMN_BG_DATE] = update.bg_date

    return updated_rows


def is_bg_material(path: str | Path) -> bool:
    return Path(path).is_file() and Path(path).suffix.casefold() in BG_FILE_EXTENSIONS


def _iter_candidate_materials(root: Path) -> list[Path]:
    if root.is_file():
        return [root]

    if extract_cut_identifiers(root.name):
        return [root]

    child_materials = sorted(
        (entry for entry in root.iterdir() if entry.is_dir() or is_bg_material(entry)),
        key=lambda entry: entry.name.casefold(),
    )
    return child_materials if child_materials else [root]


def _candidate_name(path: Path) -> str:
    if is_bg_material(path):
        return path.stem
    return path.name


def _build_material_row(
    cut_identifier: CutIdentifier,
    import_date: str,
    is_compatible: bool,
    is_bg: bool,
) -> list[str]:
    row = [""] * len(CSV_HEADERS)
    row[COLUMN_CUT_NUMBER] = cut_identifier.cut_number
    row[COLUMN_AB_GROUP] = cut_identifier.ab_group
    row[COLUMN_STATUS] = "兼用" if is_compatible else ""
    if is_bg:
        row[COLUMN_BG_LOAD_COUNT] = "1"
        row[COLUMN_BG_DATE] = import_date
    else:
        row[COLUMN_TP_LOAD_COUNT] = "1"
        row[COLUMN_TP_DATE] = import_date
    row[COLUMN_TAKE] = ""
    row[COLUMN_TAKE_NUMBER] = ""
    row[COLUMN_DELIVERY_DATE] = ""
    return row


def _parse_load_count(value: str) -> int:
    text = str(value or "").strip()
    if not text:
        return 0
    try:
        return int(text)
    except ValueError:
        return 0


def _normalize_row(row: list[str]) -> list[str]:
    normalized = [""] * len(CSV_HEADERS)
    for index in range(min(len(row), len(CSV_HEADERS))):
        normalized[index] = "" if row[index] is None else str(row[index])
    return normalized
