from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from cutmanager import __version__
from cutmanager.constants import SUPPORTED_PROJECT_EXTENSIONS
from cutmanager.main_window import MainWindow


APP_ICON_PATH = Path(__file__).resolve().parent / "assets" / "cutmanager_icon.ico"


def resolve_startup_file(argv: list[str]) -> str | None:
    """起動引数から開くファイルを決める。

    .cutmgr をエクスプローラーでダブルクリックすると、関連付け経由で
    ``CutManager.exe "<ファイルパス>"`` の形で起動する。
    """

    for argument in argv[1:]:
        if not argument or argument.startswith("-"):
            continue
        candidate = Path(argument)
        if candidate.suffix.casefold() not in SUPPORTED_PROJECT_EXTENSIONS:
            continue
        if not candidate.is_file():
            continue
        return str(candidate)
    return None


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("CutManager")
    app.setApplicationVersion(__version__)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))

    window = MainWindow()
    if APP_ICON_PATH.exists():
        window.setWindowIcon(QIcon(str(APP_ICON_PATH)))

    # 前回のファイルを復元したあとに開き直し、引数で指定されたファイルを優先する。
    startup_file = resolve_startup_file(sys.argv)
    if startup_file:
        window.open_csv_path(startup_file)

    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
