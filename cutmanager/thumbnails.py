from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import (
    QObject,
    QRunnable,
    QSize,
    QStandardPaths,
    Qt,
    QThreadPool,
    QTimer,
    QUrl,
    Signal,
)
from PySide6.QtGui import QImage, QPixmap

try:  # QtMultimedia は PySide6 に同梱されるが、環境によっては利用できない。
    from PySide6.QtMultimedia import QMediaPlayer, QVideoFrame, QVideoSink

    _MULTIMEDIA_AVAILABLE = True
except Exception:  # pragma: no cover - 環境依存
    QMediaPlayer = None  # type: ignore[assignment]
    QVideoSink = None  # type: ignore[assignment]
    QVideoFrame = None  # type: ignore[assignment]
    _MULTIMEDIA_AVAILABLE = False


# 生成する元サムネイルのサイズ。セル内へアスペクト比維持で縮小表示するため、
# セルより大きめに作っておき、行/列を広げても粗くならないようにする。
THUMBNAIL_SIZE = QSize(160, 90)
# 先頭 8 フレームはカット情報（スレート）表示のため、9 フレーム目（index 8）を採用する。
_THUMBNAIL_FRAME_INDEX = 9
# フレームが十分に届かない場合のタイムアウト（ミリ秒）: QtMultimedia フォールバック用。
_FRAME_TIMEOUT_MS = 3000
# ffmpeg 1 本あたりのタイムアウト（秒）。
_FFMPEG_TIMEOUT_S = 20
# ffmpeg を同時に走らせる本数（並列生成でバッチを高速化する）。
_MAX_PARALLEL = 4

_ffmpeg_path_cache: str | None = None
_ffmpeg_lookup_done = False


def multimedia_available() -> bool:
    return _MULTIMEDIA_AVAILABLE


def find_ffmpeg() -> str | None:
    """ffmpeg の実行ファイルを探す（PATH → 実行ファイル同梱ディレクトリの順）。"""
    global _ffmpeg_path_cache, _ffmpeg_lookup_done
    if _ffmpeg_lookup_done:
        return _ffmpeg_path_cache

    _ffmpeg_lookup_done = True
    exe_name = "ffmpeg.exe" if sys.platform.startswith("win") else "ffmpeg"

    candidate = shutil.which("ffmpeg")
    if candidate:
        _ffmpeg_path_cache = candidate
        return candidate

    # 配布物に同梱された ffmpeg を探す。
    # - Nuitka onefile: データファイルは展開先（__file__ の 1 つ上）に配置される。
    # - Nuitka standalone / 開発時: 実行ファイルや argv[0] と同じ場所。
    # - PyInstaller: sys._MEIPASS。
    search_dirs: list[Path] = []

    def _add(path: Path | None) -> None:
        if path is not None and path not in search_dirs:
            search_dirs.append(path)

    try:
        _add(Path(__file__).resolve().parent.parent)  # Nuitka のデータ展開先ルート
    except Exception:  # pragma: no cover
        pass
    try:
        _add(Path(sys.executable).resolve().parent)
    except Exception:  # pragma: no cover
        pass
    if sys.argv and sys.argv[0]:
        try:
            _add(Path(sys.argv[0]).resolve().parent)
        except Exception:  # pragma: no cover
            pass
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        _add(Path(meipass))

    relative_candidates = (
        Path(exe_name),
        Path("ffmpeg") / exe_name,
        Path("assets") / "ffmpeg" / exe_name,  # 開発時（リポジトリ内）
    )
    for directory in search_dirs:
        for relative in relative_candidates:
            bundled = directory / relative
            if bundled.is_file():
                _ffmpeg_path_cache = str(bundled)
                return _ffmpeg_path_cache

    _ffmpeg_path_cache = None
    return None


def _cache_dir() -> Path:
    base = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.CacheLocation)
    if not base:
        base = str(Path.home() / ".cutmanager_cache")
    directory = Path(base) / "thumbnails"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def _cache_key(video_path: str) -> str:
    path = Path(video_path)
    try:
        mtime = path.stat().st_mtime_ns
        size = path.stat().st_size
    except OSError:
        mtime = 0
        size = 0
    raw = f"{path.resolve(strict=False)}|{mtime}|{size}".encode("utf-8", "replace")
    return hashlib.sha1(raw).hexdigest()


def _normalize_path(video_path: str) -> str:
    # キャッシュキー用の正規化。描画のたびに呼ばれるためファイルシステムには触れず、
    # 純粋な文字列処理（絶対パス化＋大小文字正規化）のみで行う。
    text = str(video_path or "").strip()
    if not text:
        return ""
    try:
        return os.path.normcase(os.path.abspath(text))
    except (OSError, ValueError):
        return text


class _FfmpegSignals(QObject):
    finished = Signal(str, str, bool)  # key, output_path, success


class _FfmpegThumbnailTask(QRunnable):
    """ffmpeg で 9 フレーム目の 1 枚だけを抜き出してキャッシュ PNG を生成するワーカー。

    キャッシュキー算出（stat）・ディスク存在確認・生成のいずれもワーカースレッドで
    行い、GUI スレッドをファイル I/O でブロックしない。
    """

    def __init__(self, ffmpeg_path: str, video_path: str, key: str, force: bool, signals: _FfmpegSignals) -> None:
        super().__init__()
        self._ffmpeg_path = ffmpeg_path
        self._video_path = video_path
        self._key = key
        self._force = force
        self._signals = signals

    def run(self) -> None:  # QThreadPool のワーカースレッドで実行される。
        output_path = ""
        success = False
        try:
            output_path = str(_cache_dir() / f"{_cache_key(self._video_path)}.png")
            if not self._force and self._is_valid_output(output_path):
                # 既存キャッシュが使える場合は ffmpeg を起動しない。
                success = True
            else:
                success = self._generate(output_path)
        except Exception:  # pragma: no cover - 生成失敗は失敗として通知する。
            success = False
        self._signals.finished.emit(self._key, output_path if success else "", success)

    @staticmethod
    def _is_valid_output(output_path: str) -> bool:
        try:
            output = Path(output_path)
            return output.is_file() and output.stat().st_size > 0
        except OSError:
            return False

    def _generate(self, output_path: str) -> bool:
        width = THUMBNAIL_SIZE.width()
        # 先頭 8 フレームを飛ばし 9 フレーム目を選択、指定サイズに収める。
        video_filter = f"select=gte(n\\,{_THUMBNAIL_FRAME_INDEX - 1}),scale={width}:-2"
        command = [
            self._ffmpeg_path,
            "-nostdin",
            "-v",
            "error",
            "-i",
            self._video_path,
            "-vf",
            video_filter,
            "-frames:v",
            "1",
            "-y",
            output_path,
        ]
        creationflags = 0
        if sys.platform.startswith("win"):
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
            timeout=_FFMPEG_TIMEOUT_S,
            creationflags=creationflags,
        )
        return result.returncode == 0 and self._is_valid_output(output_path)


class ThumbnailProvider(QObject):
    """動画パスからサムネイルを遅延生成し、メモリ／ディスクにキャッシュする。

    - ffmpeg が使える場合はワーカースレッドで並列生成し、GUI を止めず高速。
    - ffmpeg が無い場合は QtMultimedia で逐次生成にフォールバックする。
    - 同一パス（+更新時刻/サイズ）は再生成しない。生成完了時に thumbnailReady を発火する。
    """

    thumbnailReady = Signal(str)  # video_path
    progressChanged = Signal(int, int)  # done, total（total==0 はアイドル）

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._memory_cache: dict[str, QPixmap] = {}
        self._requested: set[str] = set()
        self._failed: set[str] = set()
        # ディスクに生成済みだがまだ QPixmap を読み込んでいないキー → 出力PNGパス。
        # 読み込みは GUI スレッドで行うため、描画時（可視セルのみ）に遅延実行する。
        self._disk_ready: dict[str, str] = {}
        # QtMultimedia フォールバック時に、キーから元パスを引くための対応表。
        self._path_by_key: dict[str, str] = {}
        # パス正規化（resolve）は描画のたびに呼ばれるためメモ化する。
        self._norm_cache: dict[str, str] = {}
        # 生成進捗（実際に生成した件数のみ数える。キャッシュ命中は含めない）。
        self._gen_total = 0
        self._gen_done = 0

        self._ffmpeg_path = find_ffmpeg()
        self._pool = QThreadPool(self)
        self._pool.setMaxThreadCount(min(_MAX_PARALLEL, max(1, QThreadPool.globalInstance().maxThreadCount())))
        self._ffmpeg_signals = _FfmpegSignals(self)
        self._ffmpeg_signals.finished.connect(self._on_ffmpeg_finished)

        # QtMultimedia フォールバック用の状態。
        self._pending: list[str] = []
        self._busy = False
        self._player = None
        self._sink = None
        self._current_path: str | None = None
        self._frame_count = 0
        self._last_pixmap: QPixmap | None = None
        self._timeout_timer: QTimer | None = None

    def _norm(self, video_path: str) -> str:
        """正規化キー（ファイルシステムに触れない）をメモ化して返す（描画ホットパス用）。"""
        raw = str(video_path or "").strip()
        if not raw:
            return ""
        cached = self._norm_cache.get(raw)
        if cached is None:
            cached = _normalize_path(raw)
            self._norm_cache[raw] = cached
        return cached

    def thumbnail(self, video_path: str) -> QPixmap | None:
        """キャッシュ済みサムネイルを返す。未生成なら生成を予約して None を返す。"""
        key = self._norm(video_path)
        if not key:
            return None
        cached = self._memory_cache.get(key)
        if cached is not None:
            return cached
        # 生成済み（ディスクにある）なら、この描画タイミングで初めて読み込む。
        output_path = self._disk_ready.pop(key, None)
        if output_path is not None:
            pixmap = QPixmap(output_path)
            if not pixmap.isNull():
                self._memory_cache[key] = pixmap
                return pixmap
            self._failed.add(key)
            return None
        # まだ結論が出ていないキーだけ生成を予約する。
        if key not in self._requested and key not in self._failed:
            self._request(key, str(video_path or "").strip(), force=False)
        return None

    def request(self, video_path: str, force: bool = False) -> None:
        self._request(self._norm(video_path), str(video_path or "").strip(), force=force)

    def _request(self, key: str, video_path: str, force: bool) -> None:
        # GUI スレッドではファイルシステムに触れない（stat 等はワーカースレッドで行う）。
        if not key or not video_path:
            return
        if not force and (
            key in self._memory_cache
            or key in self._disk_ready
            or key in self._requested
            or key in self._failed
        ):
            return
        self._failed.discard(key)
        self._requested.add(key)
        self._path_by_key[key] = video_path

        if self._ffmpeg_path:
            self._begin_generation()
            self._submit_ffmpeg(key, video_path, force)
            return
        if _MULTIMEDIA_AVAILABLE:
            self._begin_generation()
            self._pending.append(key)
            QTimer.singleShot(0, self._process_next)
            return
        self._requested.discard(key)
        self._failed.add(key)

    def clear_pending(self) -> None:
        self._pending.clear()

    def invalidate(self, video_paths) -> None:
        """指定パスのメモリ状態を破棄し、再生成可能にする（ディスクI/Oはしない）。

        実際のディスク上のキャッシュ PNG は、再生成（force）時に ffmpeg が上書きする。
        """
        for video_path in video_paths:
            key = self._norm(video_path)
            if not key:
                continue
            self._memory_cache.pop(key, None)
            self._requested.discard(key)
            self._failed.discard(key)
            self._disk_ready.pop(key, None)

    # ------------------------------------------------------------------
    # 進捗
    # ------------------------------------------------------------------
    def _begin_generation(self) -> None:
        self._gen_total += 1
        self.progressChanged.emit(self._gen_done, self._gen_total)

    def _end_generation(self) -> None:
        self._gen_done += 1
        if self._gen_done >= self._gen_total:
            # バッチ完了。カウンターをリセットしてアイドル状態を通知する。
            self._gen_total = 0
            self._gen_done = 0
        self.progressChanged.emit(self._gen_done, self._gen_total)

    # ------------------------------------------------------------------
    # 高速パス: ffmpeg（並列）
    # ------------------------------------------------------------------
    def _submit_ffmpeg(self, key: str, video_path: str, force: bool) -> None:
        task = _FfmpegThumbnailTask(self._ffmpeg_path, video_path, key, force, self._ffmpeg_signals)
        self._pool.start(task)

    def _on_ffmpeg_finished(self, key: str, output_path: str, success: bool) -> None:
        self._requested.discard(key)
        # 完了通知は多発するため、ここでは QPixmap を読み込まず（重いので）
        # 「ディスク準備完了」＋出力パスだけ記録し、実際の読込は描画時に遅延させる。
        if success and output_path:
            self._disk_ready[key] = output_path
            self.thumbnailReady.emit(key)
            self._end_generation()
            return
        self._failed.add(key)
        self._end_generation()

    # ------------------------------------------------------------------
    # フォールバック: QtMultimedia（逐次）
    # ------------------------------------------------------------------
    def _ensure_player(self) -> None:
        # プレイヤー／シンクは 1 組だけ生成して使い回す（ファイルごとの生成コストを避ける）。
        if self._player is not None:
            return
        self._sink = QVideoSink(self)
        self._sink.videoFrameChanged.connect(self._on_frame)
        self._player = QMediaPlayer(self)
        self._player.setVideoSink(self._sink)
        self._player.errorOccurred.connect(lambda *_: self._finish_failure())
        self._player.mediaStatusChanged.connect(self._on_media_status)

    def _process_next(self) -> None:
        if self._busy or not self._pending:
            return
        key = self._pending.pop(0)
        if key in self._memory_cache or key in self._failed:
            self._requested.discard(key)
            QTimer.singleShot(0, self._process_next)
            return

        self._busy = True
        self._current_path = key
        self._frame_count = 0
        self._last_pixmap = None

        self._ensure_player()
        # 先頭から再生し、9 フレーム目の有効フレームを採用する（fps 非依存）。
        source_path = self._path_by_key.get(key, key)
        self._player.setSource(QUrl.fromLocalFile(source_path))
        self._player.play()
        if self._timeout_timer is None:
            self._timeout_timer = QTimer(self)
            self._timeout_timer.setSingleShot(True)
            self._timeout_timer.timeout.connect(self._on_timeout)
        self._timeout_timer.start(_FRAME_TIMEOUT_MS)

    def _on_frame(self, frame) -> None:
        if not self._busy or frame is None or not frame.isValid():
            return
        image = frame.toImage()
        if image.isNull():
            return
        pixmap = self._build_pixmap(image)
        if pixmap is None:
            return
        self._frame_count += 1
        self._last_pixmap = pixmap
        if self._frame_count >= _THUMBNAIL_FRAME_INDEX:
            self._commit(pixmap)

    def _on_media_status(self, status) -> None:
        if not self._busy or QMediaPlayer is None:
            return
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            if self._last_pixmap is not None:
                self._commit(self._last_pixmap)
            else:
                self._finish_failure()
        elif status == QMediaPlayer.MediaStatus.InvalidMedia:
            self._finish_failure()

    def _on_timeout(self) -> None:
        if self._last_pixmap is not None:
            self._commit(self._last_pixmap)
        else:
            self._finish_failure()

    def _commit(self, pixmap: QPixmap) -> None:
        key = self._current_path
        if key is None:
            return
        self._memory_cache[key] = pixmap
        self._save_to_disk(key, pixmap)
        self._requested.discard(key)
        self.thumbnailReady.emit(key)
        self._end_generation()
        self._finish_success()

    def _build_pixmap(self, image: QImage) -> QPixmap | None:
        if image.isNull():
            return None
        scaled = image.scaled(
            THUMBNAIL_SIZE,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        return QPixmap.fromImage(scaled)

    def _finish_success(self) -> None:
        if not self._busy:
            return
        self._release_current()
        self._busy = False
        QTimer.singleShot(0, self._process_next)

    def _finish_failure(self) -> None:
        if not self._busy:
            return
        if self._current_path is not None:
            self._requested.discard(self._current_path)
            self._failed.add(self._current_path)
        self._end_generation()
        self._release_current()
        self._busy = False
        QTimer.singleShot(0, self._process_next)

    def _release_current(self) -> None:
        # プレイヤーは破棄せず、現在の再生だけ止めて次に備える。
        if self._timeout_timer is not None:
            self._timeout_timer.stop()
        if self._player is not None:
            try:
                self._player.stop()
                self._player.setSource(QUrl())
            except Exception:  # pragma: no cover
                pass
        self._current_path = None
        self._frame_count = 0
        self._last_pixmap = None

    # ------------------------------------------------------------------
    # ディスクキャッシュ
    # ------------------------------------------------------------------
    def _disk_path(self, video_path: str) -> Path:
        return _cache_dir() / f"{_cache_key(video_path)}.png"

    def _save_to_disk(self, video_path: str, pixmap: QPixmap) -> None:
        try:
            pixmap.save(str(self._disk_path(video_path)), "PNG")
        except Exception:  # pragma: no cover
            pass
