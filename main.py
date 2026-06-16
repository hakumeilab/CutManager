from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from cutmanager import __version__
from cutmanager.main_window import MainWindow


APP_ICON_PATH = Path(__file__).resolve().parent / "assets" / "cutmanager_icon.ico"


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("CutManager")
    app.setApplicationVersion(__version__)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))

    window = MainWindow()
    if APP_ICON_PATH.exists():
        window.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
