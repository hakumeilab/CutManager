from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from cutmanager.csv_io import load_csv_file


class CsvIoTests(unittest.TestCase):
    def test_legacy_material_headers_map_to_tp_columns(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            csv_path = Path(temp_dir) / "cuts.csv"
            with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(["カット番号", "AB分け", "区分", "素材入れ回数", "素材入れ日", "テイク", "テイク番号", "納品日"])
                writer.writerow(["001", "", "", "2", "2026/04/15", "T", "1", "2026/04/16"])

            result = load_csv_file(str(csv_path))

        self.assertEqual(result.rows[0], ["001", "", "", "", "2", "2026/04/15", "", "", "T", "1", "2026/04/16"])


if __name__ == "__main__":
    unittest.main()
