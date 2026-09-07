from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cutmanager.video_import import apply_videos_to_rows, extract_video_metadata


class VideoImportTests(unittest.TestCase):
    def test_unmatched_count_is_tracked_per_file(self) -> None:
        rows = [["001", "", "", "", "", "", "", "", "", "", ""]]

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            unmatched_video = temp_path / "010_011_take01.mov"
            matched_video = temp_path / "001_take02.mov"
            unmatched_video.write_bytes(b"unmatched")
            matched_video.write_bytes(b"matched")

            result = apply_videos_to_rows(
                [unmatched_video, matched_video],
                rows,
                "2026/04/16",
            )

        self.assertEqual(result.updated_count, 1)
        self.assertEqual(result.unmatched_count, 1)
        self.assertEqual(result.unmatched_files, [unmatched_video.name])


    def test_two_and_four_digit_cut_numbers_are_extracted(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            two_digit = temp_path / "12_take01.mov"
            four_digit = temp_path / "1024A_T2.mov"
            two_digit.write_bytes(b"a")
            four_digit.write_bytes(b"b")

            two_metadata = extract_video_metadata(two_digit)
            four_metadata = extract_video_metadata(four_digit)

        assert two_metadata is not None
        assert four_metadata is not None
        self.assertEqual([item.cut_number for item in two_metadata.cut_identifiers], ["12"])
        self.assertEqual(two_metadata.take_number, "01")
        self.assertEqual(
            [(item.cut_number, item.ab_group) for item in four_metadata.cut_identifiers],
            [("1024", "A")],
        )
        self.assertEqual(four_metadata.take_number, "2")


if __name__ == "__main__":
    unittest.main()
