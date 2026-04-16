from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from cutmanager import update_manager


class UpdateManagerTests(unittest.TestCase):
    def test_prepare_executable_update_uses_replace_mode_when_packaged(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            current_executable = temp_path / "CutManager.exe"
            downloaded_executable = temp_path / "CutManager-new.exe"
            current_executable.write_bytes(b"current")
            downloaded_executable.write_bytes(b"new")

            with patch.object(update_manager, "can_apply_update_in_place", return_value=True), patch.object(
                update_manager.sys, "executable", str(current_executable)
            ), patch.object(update_manager.os, "getpid", return_value=1234):
                prepared = update_manager._prepare_executable_update(downloaded_executable)

            self.assertEqual(prepared.mode, "replace-exe")
            self.assertEqual(prepared.launch_program, "powershell.exe")
            self.assertTrue(prepared.launch_arguments[-1].endswith("apply_cutmanager_exe_update.ps1"))
            script_path = Path(prepared.launch_arguments[-1])
            self.assertTrue(script_path.exists())
            script_text = script_path.read_text(encoding="utf-8")
            self.assertIn("Move-Item -LiteralPath $downloadedExe -Destination $targetExe -Force", script_text)

    def test_prepare_executable_update_falls_back_to_installer_when_not_packaged(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            downloaded_executable = Path(temp_dir) / "CutManager-new.exe"
            downloaded_executable.write_bytes(b"new")

            with patch.object(update_manager, "can_apply_update_in_place", return_value=False):
                prepared = update_manager._prepare_executable_update(downloaded_executable)

            self.assertEqual(prepared.mode, "installer")
            self.assertEqual(prepared.launch_program, str(downloaded_executable))
            self.assertEqual(prepared.launch_arguments, [])

    def test_build_update_script_syncs_directories_and_removes_stale_files(self) -> None:
        script = update_manager._build_update_script(
            stage_directory=Path(r"C:\stage"),
            target_directory=Path(r"C:\target"),
            relative_executable=Path("CutManager.exe"),
            process_id=1234,
        )

        self.assertIn("function Sync-CutManagerDirectory", script)
        self.assertIn("Remove-Item -LiteralPath $_.FullName -Recurse -Force", script)
        self.assertIn("Sync-CutManagerDirectory -SourceDir $stageDir -DestinationDir $targetDir", script)


if __name__ == "__main__":
    unittest.main()
