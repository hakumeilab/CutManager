from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from cutmanager import update_manager
from cutmanager.update_manager import UpdateAsset, UpdateError


class PrepareUpdateTests(unittest.TestCase):
    def test_installer_asset_runs_silently(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            downloaded = Path(temp_dir) / "CutManager-0.3.3-windows-setup.exe"
            downloaded.write_bytes(b"installer")

            prepared = update_manager.prepare_update(downloaded)

            self.assertEqual(prepared.mode, "installer")
            self.assertEqual(prepared.launch_program, str(downloaded))
            self.assertEqual(prepared.launch_arguments, update_manager.INSTALLER_SILENT_ARGS)

    def test_non_installer_exe_runs_without_silent_args(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            downloaded = Path(temp_dir) / "CutManager-0.3.3-windows-onefile.exe"
            downloaded.write_bytes(b"portable")

            prepared = update_manager.prepare_update(downloaded)

            self.assertEqual(prepared.mode, "installer")
            self.assertEqual(prepared.launch_program, str(downloaded))
            self.assertEqual(prepared.launch_arguments, [])

    def test_non_exe_asset_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            downloaded = Path(temp_dir) / "CutManager-0.3.3-windows.zip"
            downloaded.write_bytes(b"zip")

            with self.assertRaises(UpdateError):
                update_manager.prepare_update(downloaded)


class SelectAssetTests(unittest.TestCase):
    def _asset(self, name: str, size: int = 1000) -> UpdateAsset:
        return UpdateAsset(name=name, download_url=f"https://example/{name}", size=size, content_type="")

    def test_prefers_setup_installer_over_onefile(self) -> None:
        assets = [
            self._asset("CutManager-0.3.3-windows-onefile.exe"),
            self._asset("CutManager-0.3.3-windows-setup.exe"),
            self._asset("CutManager-0.3.3-windows-onefile.sha256.txt"),
        ]

        selected = update_manager._select_release_asset(assets)

        self.assertIsNotNone(selected)
        self.assertEqual(selected.name, "CutManager-0.3.3-windows-setup.exe")

    def test_falls_back_to_onefile_when_no_installer(self) -> None:
        assets = [self._asset("CutManager-0.3.3-windows-onefile.exe")]

        selected = update_manager._select_release_asset(assets)

        self.assertIsNotNone(selected)
        self.assertEqual(selected.name, "CutManager-0.3.3-windows-onefile.exe")

    def test_returns_none_without_exe_asset(self) -> None:
        assets = [self._asset("CutManager-0.3.3-windows-onefile.sha256.txt")]

        self.assertIsNone(update_manager._select_release_asset(assets))


class ChecksumVerificationTests(unittest.TestCase):
    def test_matching_hash_passes(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "CutManager-setup.exe"
            payload = b"payload-bytes"
            target.write_bytes(payload)
            digest = hashlib.sha256(payload).hexdigest()

            asset = UpdateAsset(
                name=target.name,
                download_url="https://example/setup.exe",
                size=len(payload),
                content_type="",
                sha256_url="https://example/setup.exe.sha256.txt",
            )

            with patch.object(update_manager, "_read_text", return_value=f"{digest} *{target.name}"):
                update_manager._verify_downloaded_asset(asset, target)  # should not raise

    def test_mismatching_hash_raises(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "CutManager-setup.exe"
            target.write_bytes(b"payload-bytes")
            wrong = "0" * 64

            asset = UpdateAsset(
                name=target.name,
                download_url="https://example/setup.exe",
                size=13,
                content_type="",
                sha256_url="https://example/setup.exe.sha256.txt",
            )

            with patch.object(update_manager, "_read_text", return_value=f"{wrong} *{target.name}"):
                with self.assertRaises(UpdateError):
                    update_manager._verify_downloaded_asset(asset, target)

    def test_missing_checksum_url_skips_verification(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            target = Path(temp_dir) / "CutManager-setup.exe"
            target.write_bytes(b"payload-bytes")
            asset = UpdateAsset(name=target.name, download_url="x", size=13, content_type="", sha256_url="")

            update_manager._verify_downloaded_asset(asset, target)  # should not raise


if __name__ == "__main__":
    unittest.main()
