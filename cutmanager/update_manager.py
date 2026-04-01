from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, Signal

from . import __version__

APP_NAME = "CutManager"
GITHUB_OWNER = "hakumeilab"
GITHUB_REPOSITORY = "CutManager"
LATEST_RELEASE_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPOSITORY}/releases/latest"
RELEASES_PAGE_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPOSITORY}/releases"
HTTP_TIMEOUT_SECONDS = 20
DOWNLOAD_CHUNK_SIZE = 1024 * 128


class UpdateError(RuntimeError):
    pass


@dataclass(slots=True)
class UpdateAsset:
    name: str
    download_url: str
    size: int
    content_type: str

    @property
    def suffix(self) -> str:
        return Path(self.name).suffix.casefold()


@dataclass(slots=True)
class ReleaseInfo:
    version: str
    tag_name: str
    title: str
    body: str
    html_url: str
    published_at: str
    asset: UpdateAsset | None


@dataclass(slots=True)
class UpdateCheckResult:
    current_version: str
    release: ReleaseInfo
    update_available: bool


@dataclass(slots=True)
class PreparedUpdate:
    launch_program: str
    launch_arguments: list[str]
    mode: str
    downloaded_path: Path


class UpdateCheckWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def run(self) -> None:
        try:
            result = check_for_updates()
        except UpdateError as exc:
            self.failed.emit(str(exc))
            return
        self.finished.emit(result)


class UpdateDownloadWorker(QObject):
    progress = Signal(int, int)
    finished = Signal(str)
    failed = Signal(str)

    def __init__(self, asset: UpdateAsset) -> None:
        super().__init__()
        self._asset = asset

    def run(self) -> None:
        try:
            downloaded_path = download_release_asset(self._asset, self.progress.emit)
        except UpdateError as exc:
            self.failed.emit(str(exc))
            return
        self.finished.emit(str(downloaded_path))


def check_for_updates() -> UpdateCheckResult:
    current_version = normalize_version(__version__)
    release = fetch_latest_release()
    return UpdateCheckResult(
        current_version=current_version,
        release=release,
        update_available=is_newer_version(release.version, current_version),
    )


def fetch_latest_release() -> ReleaseInfo:
    payload = _read_json(LATEST_RELEASE_API_URL)
    tag_name = str(payload.get("tag_name") or "").strip()
    version = normalize_version(tag_name or str(payload.get("name") or "").strip())
    if not version:
        raise UpdateError("最新リリースのバージョン表記を解釈できませんでした。")

    assets = [_parse_asset(raw_asset) for raw_asset in payload.get("assets", [])]
    selected_asset = _select_release_asset(assets)

    return ReleaseInfo(
        version=version,
        tag_name=tag_name,
        title=str(payload.get("name") or tag_name or version),
        body=str(payload.get("body") or "").strip(),
        html_url=str(payload.get("html_url") or RELEASES_PAGE_URL),
        published_at=format_release_timestamp(str(payload.get("published_at") or "")),
        asset=selected_asset,
    )


def download_release_asset(
    asset: UpdateAsset,
    progress_callback: Callable[[int, int], None] | None = None,
) -> Path:
    temp_root = Path(tempfile.mkdtemp(prefix="cutmanager-update-"))
    destination = temp_root / asset.name
    request = _build_request(asset.download_url)

    try:
        with urllib.request.urlopen(request, timeout=HTTP_TIMEOUT_SECONDS) as response:
            total_bytes = int(response.headers.get("Content-Length", "0") or 0)
            downloaded_bytes = 0
            with destination.open("wb") as handle:
                while True:
                    chunk = response.read(DOWNLOAD_CHUNK_SIZE)
                    if not chunk:
                        break
                    handle.write(chunk)
                    downloaded_bytes += len(chunk)
                    if progress_callback is not None:
                        progress_callback(downloaded_bytes, total_bytes)
    except urllib.error.HTTPError as exc:
        raise UpdateError(f"更新ファイルをダウンロードできませんでした: HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise UpdateError(f"更新ファイルをダウンロードできませんでした: {exc.reason}") from exc
    except OSError as exc:
        raise UpdateError(f"更新ファイルを保存できませんでした: {exc}") from exc

    if progress_callback is not None:
        progress_callback(asset.size or destination.stat().st_size, asset.size or destination.stat().st_size)

    return destination


def prepare_update(downloaded_path: Path) -> PreparedUpdate:
    suffix = downloaded_path.suffix.casefold()

    if suffix == ".exe":
        return PreparedUpdate(
            launch_program=str(downloaded_path),
            launch_arguments=[],
            mode="installer",
            downloaded_path=downloaded_path,
        )

    if suffix != ".zip":
        raise UpdateError("対応していない更新ファイル形式です。zip または exe を使用してください。")

    if not can_apply_update_in_place():
        raise UpdateError(
            "zip 更新の自動適用は、Windows の配布版でのみ使用できます。"
        )

    current_executable = Path(sys.executable).resolve()
    target_directory = current_executable.parent
    extracted_root = _extract_update_archive(downloaded_path)
    payload_root = _resolve_payload_root(extracted_root, current_executable.name)
    relative_executable = _find_relative_executable(payload_root, current_executable.name)

    script_path = downloaded_path.parent / "apply_cutmanager_update.ps1"
    script_path.write_text(
        _build_update_script(
            stage_directory=payload_root,
            target_directory=target_directory,
            relative_executable=relative_executable,
            process_id=os.getpid(),
        ),
        encoding="utf-8",
    )

    return PreparedUpdate(
        launch_program="powershell.exe",
        launch_arguments=[
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-File",
            str(script_path),
        ],
        mode="in-place",
        downloaded_path=downloaded_path,
    )


def can_apply_update_in_place() -> bool:
    return sys.platform.startswith("win") and _is_packaged_runtime()


def _is_packaged_runtime() -> bool:
    if bool(getattr(sys, "frozen", False)):
        return True
    return "__compiled__" in globals()


def normalize_version(value: str) -> str:
    normalized = str(value or "").strip()
    normalized = re.sub(r"^[^0-9]+", "", normalized)
    return normalized


def is_newer_version(latest: str, current: str) -> bool:
    return _version_key(latest) > _version_key(current)


def human_readable_size(size: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    current = float(max(0, size))
    for unit in units:
        if current < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(current)} {unit}"
            return f"{current:.1f} {unit}"
        current /= 1024.0
    return f"{int(size)} B"


def format_release_timestamp(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return "-"

    try:
        release_datetime = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return text

    local_datetime = release_datetime.astimezone(timezone.utc).astimezone()
    return local_datetime.strftime("%Y/%m/%d %H:%M")


def _build_request(url: str) -> urllib.request.Request:
    return urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"{APP_NAME}/{__version__}",
        },
    )


def _read_json(url: str) -> dict:
    request = _build_request(url)
    try:
        with urllib.request.urlopen(request, timeout=HTTP_TIMEOUT_SECONDS) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise UpdateError(
                "最新リリースが見つかりませんでした。GitHub Releases が公開されているか確認してください。"
            ) from exc
        raise UpdateError(f"更新情報の取得に失敗しました: HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise UpdateError(f"更新情報の取得に失敗しました: {exc.reason}") from exc
    except json.JSONDecodeError as exc:
        raise UpdateError("更新情報の形式が不正です。") from exc


def _parse_asset(payload: dict) -> UpdateAsset:
    return UpdateAsset(
        name=str(payload.get("name") or ""),
        download_url=str(payload.get("browser_download_url") or ""),
        size=int(payload.get("size") or 0),
        content_type=str(payload.get("content_type") or ""),
    )


def _select_release_asset(assets: list[UpdateAsset]) -> UpdateAsset | None:
    ranked_assets: list[tuple[int, UpdateAsset]] = []
    prefer_zip = can_apply_update_in_place()

    for asset in assets:
        if not asset.name or not asset.download_url:
            continue
        score = _asset_score(asset, prefer_zip=prefer_zip)
        if score <= 0:
            continue
        ranked_assets.append((score, asset))

    if not ranked_assets:
        return None

    ranked_assets.sort(
        key=lambda item: (
            item[0],
            item[1].size,
            item[1].name.casefold(),
        ),
        reverse=True,
    )
    return ranked_assets[0][1]


def _asset_score(asset: UpdateAsset, *, prefer_zip: bool) -> int:
    name = asset.name.casefold()
    suffix = asset.suffix

    if suffix == ".zip":
        score = 200 if prefer_zip else 120
    elif suffix == ".exe":
        score = 180 if prefer_zip else 200
    else:
        return -1

    if "cutmanager" in name:
        score += 50
    if "standalone" in name or "portable" in name:
        score += 30
    if "windows" in name or "win" in name:
        score += 20
    if "debug" in name or "symbols" in name or "tests" in name:
        score -= 150
    return score


def _version_key(version: str) -> tuple[tuple[int, int | str], ...]:
    parts = re.split(r"[._+\-]", normalize_version(version))
    key: list[tuple[int, int | str]] = []
    for part in parts:
        if not part:
            continue
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.casefold()))
    return tuple(key)


def _extract_update_archive(downloaded_path: Path) -> Path:
    extract_root = downloaded_path.parent / "payload"
    extract_root.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(downloaded_path) as archive:
            _safe_extract_archive(archive, extract_root)
    except OSError as exc:
        raise UpdateError(f"更新ファイルを展開できませんでした: {exc}") from exc
    except zipfile.BadZipFile as exc:
        raise UpdateError("更新ファイルの zip が壊れています。") from exc

    return extract_root


def _safe_extract_archive(archive: zipfile.ZipFile, destination: Path) -> None:
    destination_resolved = destination.resolve()
    for member in archive.infolist():
        member_path = destination / member.filename
        resolved_path = member_path.resolve(strict=False)
        if destination_resolved not in resolved_path.parents and resolved_path != destination_resolved:
            raise UpdateError("更新ファイルに不正なパスが含まれています。")
    archive.extractall(destination)


def _resolve_payload_root(extracted_root: Path, preferred_executable_name: str) -> Path:
    preferred_path = extracted_root / preferred_executable_name
    if preferred_path.exists():
        return extracted_root

    children = [child for child in extracted_root.iterdir()]
    if len(children) == 1 and children[0].is_dir():
        child = children[0]
        if (child / preferred_executable_name).exists():
            return child

    executables = list(extracted_root.rglob("*.exe"))
    if not executables:
        raise UpdateError("更新 zip 内に実行ファイルが見つかりませんでした。")

    common_parent = executables[0].parent
    if all(executable.parent == common_parent for executable in executables):
        return common_parent
    return extracted_root


def _find_relative_executable(payload_root: Path, preferred_executable_name: str) -> Path:
    preferred_path = payload_root / preferred_executable_name
    if preferred_path.exists():
        return preferred_path.relative_to(payload_root)

    executables = sorted(payload_root.rglob("*.exe"))
    if not executables:
        raise UpdateError("更新後に起動する実行ファイルが見つかりませんでした。")

    ranked = sorted(
        executables,
        key=lambda path: (
            _executable_score(path, preferred_executable_name),
            path.name.casefold(),
        ),
        reverse=True,
    )
    return ranked[0].relative_to(payload_root)


def _executable_score(path: Path, preferred_executable_name: str) -> int:
    name = path.name.casefold()
    score = 0
    if name == preferred_executable_name.casefold():
        score += 100
    if "cutmanager" in name:
        score += 50
    if name == "main.exe":
        score += 20
    return score


def _build_update_script(
    *,
    stage_directory: Path,
    target_directory: Path,
    relative_executable: Path,
    process_id: int,
) -> str:
    stage_text = _powershell_literal(stage_directory)
    target_text = _powershell_literal(target_directory)
    relative_executable_text = _powershell_literal(relative_executable)

    return (
        "$ErrorActionPreference = 'Stop'\n"
        f"$processIdToWait = {int(process_id)}\n"
        f"$stageDir = '{stage_text}'\n"
        f"$targetDir = '{target_text}'\n"
        f"$relativeExe = '{relative_executable_text}'\n"
        "for ($i = 0; $i -lt 120; $i++) {\n"
        "    $proc = Get-Process -Id $processIdToWait -ErrorAction SilentlyContinue\n"
        "    if (-not $proc) { break }\n"
        "    Start-Sleep -Milliseconds 500\n"
        "}\n"
        "New-Item -ItemType Directory -Path $targetDir -Force | Out-Null\n"
        "Get-ChildItem -LiteralPath $stageDir -Force | ForEach-Object {\n"
        "    Copy-Item -LiteralPath $_.FullName -Destination $targetDir -Recurse -Force\n"
        "}\n"
        "$targetExe = Join-Path $targetDir $relativeExe\n"
        "Start-Sleep -Milliseconds 300\n"
        "Start-Process -FilePath $targetExe\n"
    )


def _powershell_literal(path: Path) -> str:
    return str(path).replace("'", "''")
