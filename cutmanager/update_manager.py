from __future__ import annotations

import hashlib
import json
import re
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass, field
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

# Inno Setup をサイレント実行するための引数。
# 進捗バーのみ表示し、ウィザードやメッセージボックスを出さずに更新する。
# インストーラーの [Run] は skipifsilent を付けていないため、
# 更新完了後に CutManager 自身を再起動する。
INSTALLER_SILENT_ARGS = ["/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART"]


class UpdateError(RuntimeError):
    pass


@dataclass(slots=True)
class UpdateAsset:
    name: str
    download_url: str
    size: int
    content_type: str
    sha256_url: str = ""

    @property
    def suffix(self) -> str:
        return Path(self.name).suffix.casefold()

    @property
    def is_installer(self) -> bool:
        name = self.name.casefold()
        return self.suffix == ".exe" and ("setup" in name or "installer" in name)


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
    launch_arguments: list[str] = field(default_factory=list)
    mode: str = "installer"
    downloaded_path: Path | None = None


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
    if selected_asset is not None:
        selected_asset.sha256_url = _find_checksum_url(assets, selected_asset.name)

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

    _verify_downloaded_asset(asset, destination)

    if progress_callback is not None:
        final_size = asset.size or destination.stat().st_size
        progress_callback(final_size, final_size)

    return destination


def prepare_update(downloaded_path: Path) -> PreparedUpdate:
    """ダウンロードした更新ファイルの適用方法を決める。

    設計方針: アプリが自分自身の exe を書き換えることはしない。
    更新は必ず Inno Setup 製インストーラー (setup.exe) に委譲する。
    インストーラーは per-user (localappdata) にインストールするため管理者権限は不要で、
    ファイル置換と再起動をインストーラー側が安全に行う。
    """
    suffix = downloaded_path.suffix.casefold()
    if suffix != ".exe":
        raise UpdateError("対応していない更新ファイル形式です。インストーラー (.exe) を使用してください。")

    name = downloaded_path.name.casefold()
    is_installer = "setup" in name or "installer" in name
    arguments = list(INSTALLER_SILENT_ARGS) if is_installer else []

    return PreparedUpdate(
        launch_program=str(downloaded_path),
        launch_arguments=arguments,
        mode="installer",
        downloaded_path=downloaded_path,
    )


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
        if exc.code == 403:
            reset_text = _format_rate_limit_reset_time(exc.headers.get("X-RateLimit-Reset", ""))
            retry_message = (
                f"{reset_text} 以降に再試行してください。"
                if reset_text
                else "しばらく待ってから再試行してください。"
            )
            raise UpdateError(
                "GitHub API のレート制限により更新情報を取得できませんでした。"
                f"{retry_message}\n"
                f"手動確認: {RELEASES_PAGE_URL}"
            ) from exc
        if exc.code == 404:
            raise UpdateError(
                "最新リリースが見つかりませんでした。GitHub Releases が公開されているか確認してください。"
            ) from exc
        raise UpdateError(f"更新情報の取得に失敗しました: HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise UpdateError(f"更新情報の取得に失敗しました: {exc.reason}") from exc
    except json.JSONDecodeError as exc:
        raise UpdateError("更新情報の形式が不正です。") from exc


def _read_text(url: str) -> str:
    request = _build_request(url)
    try:
        with urllib.request.urlopen(request, timeout=HTTP_TIMEOUT_SECONDS) as response:
            return response.read().decode("utf-8", errors="replace")
    except (urllib.error.URLError, OSError):
        return ""


def _parse_asset(payload: dict) -> UpdateAsset:
    return UpdateAsset(
        name=str(payload.get("name") or ""),
        download_url=str(payload.get("browser_download_url") or ""),
        size=int(payload.get("size") or 0),
        content_type=str(payload.get("content_type") or ""),
    )


def _format_rate_limit_reset_time(value: str) -> str:
    text = str(value or "").strip()
    if not text.isdigit():
        return ""

    try:
        reset_datetime = datetime.fromtimestamp(int(text), tz=timezone.utc).astimezone()
    except (OSError, ValueError, OverflowError):
        return ""

    return reset_datetime.strftime("%Y/%m/%d %H:%M")


def _select_release_asset(assets: list[UpdateAsset]) -> UpdateAsset | None:
    """更新に使う exe を選ぶ。setup 版インストーラーを最優先する。"""
    candidates = [
        asset
        for asset in assets
        if asset.name and asset.download_url and asset.suffix == ".exe"
    ]
    if not candidates:
        return None

    candidates.sort(
        key=lambda asset: (_asset_score(asset), asset.size, asset.name.casefold()),
        reverse=True,
    )
    top = candidates[0]
    if _asset_score(top) <= _MIN_ASSET_SCORE:
        return None
    return top


_MIN_ASSET_SCORE = -100


def _asset_score(asset: UpdateAsset) -> int:
    name = asset.name.casefold()
    score = 0
    if "setup" in name or "installer" in name:
        score += 100
    if "cutmanager" in name:
        score += 20
    if "onefile" in name or "portable" in name or "standalone" in name:
        score += 5
    if "windows" in name or "win" in name:
        score += 5
    if "debug" in name or "symbols" in name or "tests" in name:
        score -= 200
    return score


def _find_checksum_url(assets: list[UpdateAsset], asset_name: str) -> str:
    expected = f"{asset_name}.sha256.txt".casefold()
    for asset in assets:
        if asset.name.casefold() == expected and asset.download_url:
            return asset.download_url
    return ""


def _verify_downloaded_asset(asset: UpdateAsset, destination: Path) -> None:
    """sha256 チェックサムが公開されていれば検証する。

    チェックサムを取得できない場合は検証をスキップする（更新自体は阻害しない）。
    ハッシュが公開されていて不一致の場合のみ、改ざんの可能性として失敗させる。
    """
    if not asset.sha256_url:
        return

    checksum_text = _read_text(asset.sha256_url)
    expected_hash = _parse_sha256_text(checksum_text)
    if not expected_hash:
        return

    actual_hash = _compute_sha256(destination)
    if actual_hash != expected_hash:
        raise UpdateError(
            "ダウンロードした更新ファイルの整合性チェックに失敗しました。\n"
            "ネットワークの問題またはファイル破損の可能性があります。再試行してください。"
        )


def _parse_sha256_text(text: str) -> str:
    for line in text.splitlines():
        match = re.search(r"\b[0-9a-fA-F]{64}\b", line)
        if match:
            return match.group(0).casefold()
    return ""


def _compute_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(DOWNLOAD_CHUNK_SIZE), b""):
            digest.update(chunk)
    return digest.hexdigest().casefold()


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
