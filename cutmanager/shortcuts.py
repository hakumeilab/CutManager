from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import QSettings
from PySide6.QtGui import QKeySequence


@dataclass(frozen=True)
class ShortcutSpec:
    key: str
    label: str
    defaults: tuple[str, ...]


# 環境設定でキー割り当てを変更できる操作の一覧。順序はダイアログの表示順。
# defaults は複数指定でき、リドゥのように複数キーで発火させたい操作に使う。
SHORTCUT_SPECS: tuple[ShortcutSpec, ...] = (
    ShortcutSpec("new", "新規作成", ("Ctrl+N",)),
    ShortcutSpec("open", "開く", ("Ctrl+O",)),
    ShortcutSpec("save", "上書き保存", ("Ctrl+S",)),
    ShortcutSpec("save_as", "名前を付けて保存", ("Ctrl+Shift+S",)),
    ShortcutSpec("undo", "元に戻す", ("Ctrl+Z",)),
    ShortcutSpec("redo", "やり直し", ("Ctrl+Y", "Ctrl+Shift+Z")),
    ShortcutSpec("copy", "コピー", ("Ctrl+C",)),
    ShortcutSpec("paste", "貼り付け", ("Ctrl+V",)),
    ShortcutSpec("add_row", "行追加", ("Ins",)),
    ShortcutSpec("delete_row", "行削除", ("Ctrl+Del",)),
)

SPEC_BY_KEY: dict[str, ShortcutSpec] = {spec.key: spec for spec in SHORTCUT_SPECS}

_SETTINGS_GROUP = "shortcuts"
_SEQUENCE_SEPARATOR = "|"


def _normalize_sequences(values) -> list[str]:
    """任意の入力を正規化した（重複なしの）ショートカット文字列リストにする。"""
    if values is None:
        items: list = []
    elif isinstance(values, str):
        items = values.split(_SEQUENCE_SEPARATOR)
    elif isinstance(values, (list, tuple)):
        items = list(values)
    else:
        items = [values]

    normalized: list[str] = []
    for item in items:
        # QKeySequence を経由して "ctrl+z" などの表記ゆれを標準表記へ揃える。
        sequence = item if isinstance(item, QKeySequence) else QKeySequence(str(item).strip())
        if sequence.isEmpty():
            continue
        text = sequence.toString(QKeySequence.SequenceFormat.PortableText)
        if text and text not in normalized:
            normalized.append(text)
    return normalized


class ShortcutManager:
    """ショートカットの既定値とユーザー上書きを QSettings で永続化する。"""

    def __init__(self, settings: QSettings) -> None:
        self._settings = settings
        self._overrides: dict[str, list[str]] = {}
        self._load()

    def _load(self) -> None:
        self._settings.beginGroup(_SETTINGS_GROUP)
        try:
            for key in self._settings.childKeys():
                if key not in SPEC_BY_KEY:
                    continue
                self._overrides[key] = _normalize_sequences(self._settings.value(key))
        finally:
            self._settings.endGroup()

    def sequences(self, key: str) -> list[str]:
        spec = SPEC_BY_KEY.get(key)
        if spec is None:
            return []
        if key in self._overrides:
            return list(self._overrides[key])
        return list(spec.defaults)

    def key_sequences(self, key: str) -> list[QKeySequence]:
        result: list[QKeySequence] = []
        for text in self.sequences(key):
            sequence = QKeySequence(text)
            if not sequence.isEmpty():
                result.append(sequence)
        return result

    def all_sequences(self) -> dict[str, list[str]]:
        return {spec.key: self.sequences(spec.key) for spec in SHORTCUT_SPECS}

    def defaults(self, key: str) -> list[str]:
        spec = SPEC_BY_KEY.get(key)
        return list(spec.defaults) if spec is not None else []

    def update(self, values: dict[str, list[str]]) -> None:
        self._settings.beginGroup(_SETTINGS_GROUP)
        try:
            for spec in SHORTCUT_SPECS:
                sequences = _normalize_sequences(values.get(spec.key, []))
                if sequences == list(spec.defaults):
                    # 既定と同じなら上書きを保存せず、既定に追従させる。
                    self._overrides.pop(spec.key, None)
                    self._settings.remove(spec.key)
                else:
                    self._overrides[spec.key] = sequences
                    self._settings.setValue(spec.key, _SEQUENCE_SEPARATOR.join(sequences))
        finally:
            self._settings.endGroup()
        self._settings.sync()
