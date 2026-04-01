from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .constants import COLUMN_STATUS, CSV_HEADERS, LEGACY_STATUS_HEADERS


class CsvLoadError(Exception):
    pass


@dataclass(slots=True)
class CsvLoadResult:
    rows: list[list[str]]
    warnings: list[str]


def _normalize_row(values: list[str]) -> list[str]:
    normalized: list[str] = []
    for index in range(len(CSV_HEADERS)):
        normalized.append("" if index >= len(values) or values[index] is None else str(values[index]))
    return normalized


def load_csv_file(path: str) -> CsvLoadResult:
    csv_path = Path(path)
    warnings: list[str] = []
    normalized_rows: list[list[str]] = []

    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.reader(handle)
            source_header = next(reader, None)
            if source_header is None:
                raise CsvLoadError("CSV が空です。1 行目にヘッダーが必要です。")
            source_header = [str(cell).strip() for cell in source_header]
            if not any(source_header):
                raise CsvLoadError("1 行目のヘッダーが空です。")

            header_map = {name: index for index, name in enumerate(source_header) if name}
            column_indexes = [header_map.get(header) for header in CSV_HEADERS]
            legacy_status_indexes = [header_map.get(header) for header in LEGACY_STATUS_HEADERS]

            missing_headers = []
            for index, header in enumerate(CSV_HEADERS):
                if column_indexes[index] is not None:
                    continue
                if index == COLUMN_STATUS and any(legacy_index is not None for legacy_index in legacy_status_indexes):
                    continue
                missing_headers.append(header)
            if missing_headers:
                warnings.append(f"必須ヘッダーが不足しています: {', '.join(missing_headers)}")

            if source_header != CSV_HEADERS:
                warnings.append("ヘッダーの順番または列構成が標準と異なります。保存時に標準列へ正規化します。")

            for source_row in reader:
                normalized_rows.append(_map_source_row(source_row, column_indexes, legacy_status_indexes))
    except OSError as exc:
        raise CsvLoadError(f"CSV を開けませんでした: {exc}") from exc
    except csv.Error as exc:
        raise CsvLoadError(f"CSV の読み込みに失敗しました: {exc}") from exc

    return CsvLoadResult(rows=normalized_rows, warnings=warnings)


def save_csv_file(path: str, rows: list[list[str]]) -> None:
    csv_path = Path(path)
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=",")
            writer.writerow(CSV_HEADERS)
            for row in rows:
                writer.writerow(_normalize_row(row))
    except OSError as exc:
        raise CsvLoadError(f"CSV を保存できませんでした: {exc}") from exc
    except csv.Error as exc:
        raise CsvLoadError(f"CSV の保存に失敗しました: {exc}") from exc


def _map_source_row(
    source_row: list[str],
    column_indexes: list[int | None],
    legacy_status_indexes: list[int | None],
) -> list[str]:
    normalized_row = [
        ""
        if column_index is None or column_index >= len(source_row) or source_row[column_index] is None
        else str(source_row[column_index])
        for column_index in column_indexes
    ]
    if not normalized_row[COLUMN_STATUS]:
        normalized_row[COLUMN_STATUS] = _resolve_legacy_status(source_row, legacy_status_indexes)
    return normalized_row


def _resolve_legacy_status(source_row: list[str], legacy_status_indexes: list[int | None]) -> str:
    for header, column_index in zip(LEGACY_STATUS_HEADERS, legacy_status_indexes):
        if column_index is None or column_index >= len(source_row):
            continue
        value = source_row[column_index]
        if value is None:
            continue
        if str(value).strip():
            return header
    return ""
