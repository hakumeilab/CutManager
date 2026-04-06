from __future__ import annotations

import json
import re
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
VERSION_FILE = ROOT_DIR / "cutmanager" / "__init__.py"
CHANGELOG_FILE = ROOT_DIR / "CHANGELOG.md"
OUTPUT_FILE = ROOT_DIR / "docs" / "release.json"
RELEASES_PAGE_URL = "https://github.com/hakumeilab/CutManager/releases/latest"


def read_version() -> str:
    content = VERSION_FILE.read_text(encoding="utf-8")
    match = re.search(r'__version__\s*=\s*"([^"]+)"', content)
    if not match:
        raise RuntimeError("Version was not found in cutmanager/__init__.py.")
    return match.group(1).strip()


def read_changelog_entry(version: str) -> tuple[str, list[str]]:
    content = CHANGELOG_FILE.read_text(encoding="utf-8")
    header_pattern = re.compile(
        rf"^##\s+{re.escape(version)}(?:\s+-\s+(?P<date>\d{{4}}-\d{{2}}-\d{{2}}))?\s*$",
        re.MULTILINE,
    )
    header_match = header_pattern.search(content)
    if not header_match:
        raise RuntimeError(f"Version {version} was not found in CHANGELOG.md.")

    start = header_match.end()
    next_header_match = re.search(r"^##\s+", content[start:], re.MULTILINE)
    end = start + next_header_match.start() if next_header_match else len(content)
    notes_block = content[start:end].strip()
    if not notes_block:
        raise RuntimeError(f"CHANGELOG.md entry for version {version} is empty.")

    notes = [
        line.removeprefix("- ").strip()
        for line in notes_block.splitlines()
        if line.strip().startswith("- ")
    ]
    return header_match.group("date") or "", notes


def main() -> int:
    version = read_version()
    published_at, notes = read_changelog_entry(version)

    payload = {
        "version": version,
        "tag_name": f"v{version}",
        "published_at": published_at,
        "html_url": RELEASES_PAGE_URL,
        "download_url": RELEASES_PAGE_URL,
        "notes": notes[:3],
    }

    OUTPUT_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
