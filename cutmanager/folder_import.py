from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from .constants import (
    COLUMN_AB_GROUP,
    COLUMN_CUT_NUMBER,
    COLUMN_DELIVERY_DATE,
    COLUMN_MATERIAL_DATE,
    COLUMN_MATERIAL_LOAD_COUNT,
    COLUMN_STATUS,
    COLUMN_TAKE,
    COLUMN_TAKE_NUMBER,
    CSV_HEADERS,
)


CUT_NUMBER_PATTERN = re.compile(r"(?<!\d)(\d{3})(?!\d)")
CUT_IDENTIFIER_PATTERN = re.compile(r"(?<!\d)(\d{3})([A-Za-z]?)(?![A-Za-z0-9])")


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
    material_load_increment: int
    material_date: str

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
        if not root.is_dir():
            raise ValueError(f"フォルダーが存在しません: {root}")

        for candidate in _iter_candidate_folders(root):
            folder_key = str(candidate.resolve(strict=False)).casefold()
            if folder_key in seen_folders:
                continue
            seen_folders.add(folder_key)

            cut_identifiers = extract_cut_identifiers(candidate.name)
            if not cut_identifiers:
                failed_count += 1
                continue

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
                            material_load_increment=1,
                            material_date=import_date,
                        )
                    else:
                        update.mark_compatible = update.mark_compatible or is_compatible
                        update.material_load_increment += 1
                        update.material_date = import_date
                    continue

                row = rows_by_cut.get(cut_key)
                if row is None:
                    rows_by_cut[cut_key] = _build_material_row(cut_identifier, import_date, is_compatible)
                    continue

                row[COLUMN_STATUS] = "兼用" if is_compatible else row[COLUMN_STATUS]
                row[COLUMN_MATERIAL_LOAD_COUNT] = str(_parse_material_load_count(row[COLUMN_MATERIAL_LOAD_COUNT]) + 1)
                row[COLUMN_MATERIAL_DATE] = import_date

    return FolderImportResult(
        rows=list(rows_by_cut.values()),
        added_count=len(rows_by_cut),
        updated_count=len(updates_by_cut),
        failed_count=failed_count,
        updates=list(updates_by_cut.values()),
    )


def apply_material_updates(rows: list[list[str]], updates: list[MaterialRowUpdate]) -> list[list[str]]:
    if not updates:
        return [row.copy() for row in rows]

    updated_rows = [row.copy() for row in rows]
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
        row[COLUMN_MATERIAL_LOAD_COUNT] = str(
            _parse_material_load_count(row[COLUMN_MATERIAL_LOAD_COUNT]) + update.material_load_increment
        )
        row[COLUMN_MATERIAL_DATE] = update.material_date

    return updated_rows


def _iter_candidate_folders(root: Path) -> list[Path]:
    if extract_cut_identifiers(root.name):
        return [root]

    child_folders = sorted(
        (entry for entry in root.iterdir() if entry.is_dir()),
        key=lambda entry: entry.name.casefold(),
    )
    return child_folders if child_folders else [root]


def _build_material_row(cut_identifier: CutIdentifier, import_date: str, is_compatible: bool) -> list[str]:
    row = [""] * len(CSV_HEADERS)
    row[COLUMN_CUT_NUMBER] = cut_identifier.cut_number
    row[COLUMN_AB_GROUP] = cut_identifier.ab_group
    row[COLUMN_STATUS] = "兼用" if is_compatible else ""
    row[COLUMN_MATERIAL_LOAD_COUNT] = "1"
    row[COLUMN_MATERIAL_DATE] = import_date
    row[COLUMN_TAKE] = ""
    row[COLUMN_TAKE_NUMBER] = ""
    row[COLUMN_DELIVERY_DATE] = ""
    return row


def _parse_material_load_count(value: str) -> int:
    text = str(value or "").strip()
    if not text:
        return 0
    try:
        return int(text)
    except ValueError:
        return 0
