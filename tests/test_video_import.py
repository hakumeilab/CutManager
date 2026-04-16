from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from cutmanager.video_import import apply_videos_to_rows


class VideoImportTests(unittest.TestCase):
    def test_unmatched_count_is_tracked_per_file(self) -> None:
        rows = [["001", "", "", "", "", "", "", ""]]

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


if __name__ == "__main__":
    unittest.main()
