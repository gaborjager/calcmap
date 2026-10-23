#!/usr/bin/env python3
"""Prints the version of the Basic code, or the CHANGELOG.md section of a version.

Usage:
    python3 tools/release_notes.py version        # e.g. 0.20.0
    python3 tools/release_notes.py notes 0.20.0   # release description

Used by .github/workflows/release.yml. Fails if the version has no section
in CHANGELOG.md, so that a release is never created without a description.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def version():
    basic = (ROOT / "src" / "basic" / "CalcMap.bas").read_text(encoding="utf-8")
    return re.search(r'Const VERSION = "([\d.]+)"', basic).group(1)


def notes(v):
    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    m = re.search(rf"^## {re.escape(v)}\b.*?\n(.*?)(?=^## |\Z)", text, re.S | re.M)
    body = m.group(1).strip() if m else ""
    if not body:
        sys.exit(f"CHANGELOG.md has no section for version {v}")
    return body


if __name__ == "__main__":
    if sys.argv[1:2] == ["version"]:
        print(version())
    elif sys.argv[1:2] == ["notes"] and len(sys.argv) == 3:
        print(notes(sys.argv[2]))
    else:
        sys.exit(__doc__)
