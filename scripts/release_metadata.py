from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
VERSION_FILE = ROOT_DIR / "cutmanager" / "__init__.py"
CHANGELOG_FILE = ROOT_DIR / "CHANGELOG.md"


def read_version() -> str:
    content = VERSION_FILE.read_text(encoding="utf-8")
    match = re.search(r'__version__\s*=\s*"([^"]+)"', content)
    if not match:
        raise RuntimeError("Version was not found in cutmanager/__init__.py.")
    return match.group(1)


def normalize_version(value: str) -> str:
    return value.strip().removeprefix("refs/tags/").removeprefix("v")


def read_release_notes(version: str) -> str:
    content = CHANGELOG_FILE.read_text(encoding="utf-8")
    header_pattern = re.compile(
        rf"^##\s+{re.escape(version)}(?:\s+-\s+.+)?\s*$",
        re.MULTILINE,
    )
    header_match = header_pattern.search(content)
    if not header_match:
        raise RuntimeError(f"Version {version} was not found in CHANGELOG.md.")

    start = header_match.end()
    next_header_match = re.search(r"^##\s+", content[start:], re.MULTILINE)
    end = start + next_header_match.start() if next_header_match else len(content)
    notes = content[start:end].strip()
    if not notes:
        raise RuntimeError(f"CHANGELOG.md entry for version {version} is empty.")
    return notes


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        raise RuntimeError("Usage: release_metadata.py <version|notes> [value]")

    command = argv[1]

    if command == "version":
        print(read_version())
        return 0

    if command == "notes":
        if len(argv) < 3:
            raise RuntimeError("Usage: release_metadata.py notes <version-or-tag>")
        print(read_release_notes(normalize_version(argv[2])))
        return 0

    raise RuntimeError(f"Unsupported command: {command}")


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv))
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1) from exc
