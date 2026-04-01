from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from cutmanager.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("CutManager")

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
