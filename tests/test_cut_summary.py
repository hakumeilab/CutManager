from __future__ import annotations

import unittest

from cutmanager.main_window import calculate_cut_summary


class CutSummaryTests(unittest.TestCase):
    def test_special_count_values_exclude_unneeded_tp_and_bg(self) -> None:
        rows = [
            ["001", "", "", "", "", "", "", "", "", ""],
            ["002", "", "", "BGOnly", "", "1", "2026/06/16", "", "", ""],
            ["003", "", "", "1", "2026/06/16", "全セル", "", "", "", "2026/06/16"],
            ["004", "", "BANK", "", "", "", "", "", "", ""],
            ["005", "", "欠番", "", "", "", "", "", "", ""],
            ["006", "", "兼用", "1", "", "1", "", "", "", ""],
        ]

        summary = calculate_cut_summary(rows)

        self.assertEqual(summary["total_cuts"], 5)
        self.assertEqual(summary["delivered"], 1)
        self.assertEqual(summary["remaining_delivery"], 3)
        self.assertEqual(summary["total_tp"], 3)
        self.assertEqual(summary["tp_done"], 2)
        self.assertEqual(summary["remaining_tp"], 1)
        self.assertEqual(summary["total_bg"], 3)
        self.assertEqual(summary["bg_done"], 2)
        self.assertEqual(summary["remaining_bg"], 1)
        self.assertEqual(summary["shared"], 1)
        self.assertEqual(summary["bank"], 1)
        self.assertEqual(summary["missing"], 1)


if __name__ == "__main__":
    unittest.main()
