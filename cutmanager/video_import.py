from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from .constants import (
    COLUMN_AB_GROUP,
    COLUMN_CUT_NUMBER,
    COLUMN_DELIVERY_DATE,
    COLUMN_STATUS,
    COLUMN_TAKE,
    COLUMN_TAKE_NUMBER,
    VIDEO_FILE_EXTENSIONS,
)
from .folder_import import CUT_IDENTIFIER_PATTERN, CutIdentifier, extract_cut_identifiers, make_cut_key


EXPLICIT_TAKE_PATTERN = re.compile(r"(?i)(?:^|[_\-\s])(take|tk|t)[_\-\s]*([0-9]{1,3})(?=$|[_\-\s])")
NUMBER_GROUP_PATTERN = re.compile(r"(?<!\d)(\d{1,3})(?!\d)")


@dataclass(slots=True)
class VideoMetadata:
    cut_identifiers: list[CutIdentifier]
    compatible_label: str
    take_label: str
    take_number: str


@dataclass(slots=True)
class VideoImportResult:
    rows: list[list[str]]
    updated_count: int
    unmatched_count: int
    failed_count: int


def apply_videos_to_rows(
    video_paths: list[str | Path],
    rows: list[list[str]],
    delivery_date: str,
) -> VideoImportResult:
    updated_rows = [row.copy() for row in rows]
    row_by_cut = _build_row_map(updated_rows)
    updated_count = 0
    unmatched_count = 0
    failed_count = 0

    for video_path in sorted((Path(path) for path in video_paths), key=_video_sort_key):
        metadata = extract_video_metadata(video_path)
        if metadata is None:
            failed_count += 1
            continue

        for cut_identifier in metadata.cut_identifiers:
            row_index = row_by_cut.get(cut_identifier.key)
            if row_index is None:
                unmatched_count += 1
                continue

            row = updated_rows[row_index]
            if metadata.compatible_label:
                row[COLUMN_STATUS] = metadata.compatible_label
            row[COLUMN_TAKE] = metadata.take_label
            row[COLUMN_TAKE_NUMBER] = metadata.take_number
            row[COLUMN_DELIVERY_DATE] = delivery_date
            updated_count += 1

    return VideoImportResult(
        rows=updated_rows,
        updated_count=updated_count,
        unmatched_count=unmatched_count,
        failed_count=failed_count,
    )


def is_video_file(path: str | Path) -> bool:
    return Path(path).suffix.casefold() in VIDEO_FILE_EXTENSIONS


def extract_video_metadata(video_path: str | Path) -> VideoMetadata | None:
    path = Path(video_path)
    if path.suffix.casefold() not in VIDEO_FILE_EXTENSIONS:
        return None

    stem = path.stem
    cut_identifiers = extract_cut_identifiers(stem)
    if not cut_identifiers:
        return None

    take_number = _extract_take_number(stem, cut_identifiers)
    take_label = "T" if take_number else stem
    return VideoMetadata(
        cut_identifiers=cut_identifiers,
        compatible_label="兼用" if len(cut_identifiers) > 1 else "",
        take_label=take_label,
        take_number=take_number,
    )


def _build_row_map(rows: list[list[str]]) -> dict[tuple[str, str], int]:
    row_by_cut: dict[tuple[str, str], int] = {}
    for row_index, row in enumerate(rows):
        cut_key = make_cut_key(row[COLUMN_CUT_NUMBER], row[COLUMN_AB_GROUP])
        if cut_key[0] and cut_key not in row_by_cut:
            row_by_cut[cut_key] = row_index
    return row_by_cut


def _video_sort_key(path: Path) -> tuple[int, str]:
    try:
        modified_time = path.stat().st_mtime_ns
    except OSError:
        modified_time = -1
    return (modified_time, path.name.casefold())


def _extract_take_number(stem: str, cut_identifiers: list[CutIdentifier]) -> str:
    explicit_match = EXPLICIT_TAKE_PATTERN.search(stem)
    if explicit_match is not None:
        return explicit_match.group(2)

    last_cut_end = max(
        (match.end() for match in CUT_IDENTIFIER_PATTERN.finditer(stem)),
        default=0,
    )
    known_cut_numbers = {cut_identifier.cut_number for cut_identifier in cut_identifiers}
    fallback_numbers = [
        match.group(1)
        for match in NUMBER_GROUP_PATTERN.finditer(stem)
        if match.start() >= last_cut_end and match.group(1) not in known_cut_numbers
    ]
    if fallback_numbers:
        return fallback_numbers[-1]
    return ""
