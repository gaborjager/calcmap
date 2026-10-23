#!/usr/bin/env python3
"""Builds the CalcMap.oxt LibreOffice extension from src/.

Usage: python3 build.py
Result: dist/CalcMap.oxt

The Basic code is maintained in src/basic/CalcMap.bas (readable, diffable);
this script converts it to the .xba format expected by LibreOffice.
"""
import html
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
DIST = ROOT / "dist"

XBA_HEADER = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
    '<script:module xmlns:script="http://openoffice.org/2000/script" '
    'script:name="CalcMap" script:language="StarBasic">'
)

FILES = [
    "META-INF/manifest.xml",
    "description.xml",
    "Addons.xcu",
    "Events.xcu",
    "CalcMap/script.xlb",
    "CalcMap/dialog.xlb",
    "icons/map_16.png",
    "icons/map_26.png",
    "icons/nopreview.svg",
    "icons/codes-ok.svg",
    "icons/codes-attention.svg",
    "palettes.json",
    "CalcMap.xcs",
]
# Folders included as a whole: adding a language = adding a file
FOLDERS = ["i18n", "description"]


def main():
    basic = (SRC / "basic" / "CalcMap.bas").read_text(encoding="utf-8")

    # The version in description.xml must follow the one in the Basic code
    m = re.search(r'Const VERSION = "([\d.]+)"', basic)
    version = m.group(1) if m else "0.0.0"
    desc = (SRC / "description.xml").read_text(encoding="utf-8")
    desc = re.sub(r'<version value="[^"]*"/>', f'<version value="{version}"/>', desc)
    (SRC / "description.xml").write_text(desc, encoding="utf-8")

    xba = XBA_HEADER + html.escape(basic, quote=False) + "</script:module>\n"
    (SRC / "CalcMap" / "CalcMap.xba").write_text(xba, encoding="utf-8")

    DIST.mkdir(exist_ok=True)
    output = DIST / "CalcMap.oxt"
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        extras = [p.relative_to(SRC).as_posix() for d in FOLDERS
                  for p in sorted((SRC / d).iterdir()) if p.is_file()]
        for f in FILES + extras + ["CalcMap/CalcMap.xba"]:
            z.write(SRC / f, f)
        # Aggregated catalog of the project: fallback if the configured one is unreachable
        z.write(ROOT / "catalogue" / "catalogue.json", "catalog.json")
    print(f"{output.relative_to(ROOT)} (version {version})")


if __name__ == "__main__":
    main()
