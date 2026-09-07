from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from PySide6.QtCore import QSettings

from cutmanager.shortcuts import SPEC_BY_KEY, ShortcutManager


def _settings(tmp_dir: str) -> QSettings:
    path = str(Path(tmp_dir) / "settings.ini")
    return QSettings(path, QSettings.Format.IniFormat)


class ShortcutManagerTests(unittest.TestCase):
    def test_defaults_are_used_when_no_override(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            manager = ShortcutManager(_settings(tmp_dir))
            self.assertEqual(manager.sequences("undo"), ["Ctrl+Z"])

    def test_redo_default_includes_ctrl_shift_z(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            manager = ShortcutManager(_settings(tmp_dir))
            self.assertEqual(manager.sequences("redo"), ["Ctrl+Y", "Ctrl+Shift+Z"])
            self.assertEqual(len(manager.key_sequences("redo")), 2)

    def test_override_persists_across_instances(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            settings = _settings(tmp_dir)
            manager = ShortcutManager(settings)
            manager.update({**manager.all_sequences(), "save": ["Ctrl+Shift+K"]})

            reloaded = ShortcutManager(_settings(tmp_dir))
            self.assertEqual(reloaded.sequences("save"), ["Ctrl+Shift+K"])

    def test_setting_to_default_removes_override(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            settings = _settings(tmp_dir)
            manager = ShortcutManager(settings)
            manager.update({**manager.all_sequences(), "save": ["Ctrl+Shift+K"]})
            defaults = list(SPEC_BY_KEY["save"].defaults)
            manager.update({**manager.all_sequences(), "save": defaults})

            settings.beginGroup("shortcuts")
            stored_keys = settings.childKeys()
            settings.endGroup()
            self.assertNotIn("save", stored_keys)
            self.assertEqual(manager.sequences("save"), defaults)


if __name__ == "__main__":
    unittest.main()
