from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cutmanager.constants import COLUMN_DELIVERY_DATE, COLUMN_ROLL, COLUMN_VIDEO_PATH
from cutmanager.folder_import import extract_delivery_date, extract_roll
from cutmanager.video_import import apply_videos_to_rows


class FolderNamingTests(unittest.TestCase):
    def test_extract_roll_from_folder_name(self) -> None:
        self.assertEqual(extract_roll("BMUM_roll01_260721"), "roll01")
        self.assertEqual(extract_roll("Roll_12_final"), "roll12")
        self.assertEqual(extract_roll("ROLL003"), "roll003")
        self.assertEqual(extract_roll("no_roll_here_only_text"), "")  # 数字が続かない場合は空
        self.assertEqual(extract_roll("素材フォルダー"), "")

    def test_extract_delivery_date_yymmdd(self) -> None:
        self.assertEqual(extract_delivery_date("BMUM_roll01_260721"), "2026/07/21")
        self.assertEqual(extract_delivery_date("cut_991231_final"), "1999/12/31")
        self.assertEqual(extract_delivery_date("no_date"), "")
        # 不正な月日は無視する。
        self.assertEqual(extract_delivery_date("bad_269900"), "")


class VideoImportNamingTests(unittest.TestCase):
    def test_roll_and_path_and_date_are_applied(self) -> None:
        rows = [["001", "", "", "", "", "", "", "", "", "", "", "", "", ""]]

        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir) / "BMUM_roll01_260721"
            folder.mkdir()
            video = folder / "001_take02.mov"
            video.write_bytes(b"video")

            result = apply_videos_to_rows([video], rows, "2099/01/01")

        row = result.rows[0]
        self.assertEqual(result.updated_count, 1)
        self.assertEqual(row[COLUMN_ROLL], "roll01")
        self.assertEqual(row[COLUMN_DELIVERY_DATE], "2026/07/21")  # フォルダー名の日付を優先
        self.assertTrue(row[COLUMN_VIDEO_PATH].endswith("001_take02.mov"))

    def test_delivery_date_falls_back_to_today_when_no_folder_date(self) -> None:
        rows = [["001", "", "", "", "", "", "", "", "", "", "", "", "", ""]]

        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir) / "plain_folder"
            folder.mkdir()
            video = folder / "001_take01.mov"
            video.write_bytes(b"video")

            result = apply_videos_to_rows([video], rows, "2026/05/05")

        self.assertEqual(result.rows[0][COLUMN_DELIVERY_DATE], "2026/05/05")


if __name__ == "__main__":
    unittest.main()
