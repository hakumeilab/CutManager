"""テスト全体で QApplication を 1 つだけ先に作る。

QCoreApplication を使うテストと QApplication を使うテストを同じプロセスで
続けて実行すると、後から QApplication を作れずに異常終了するため。
"""

from __future__ import annotations

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication  # noqa: E402

if QApplication.instance() is None:
    _APP = QApplication([])
