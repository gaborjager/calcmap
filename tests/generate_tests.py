#!/usr/bin/env python3
"""Generates one test spreadsheet (.ods) per base map of the ODG Atlas catalog.

Usage:
    python3 tests/generate_tests.py                 # all the maps of the catalog
    python3 tests/generate_tests.py departements    # chosen maps (catalog id)
    python3 tests/generate_tests.py --force         # also rewrites existing files

The catalog and the maps are read from the local copy of the ODG Atlas next
to this repository (../odg-atlas, or the path in the ODG_ATLAS environment
variable), as the headless tests; without it, they are downloaded.

Each spreadsheet contains: Code | Value (test) | Name, one row per zone of the
map. Codes and names are read from the map itself (draw:name and
alternative text svg:desc); values are random but reproducible.
Codes are written as text to keep leading zeros ("01").
"""
import io
import json
import os
import random
import sys
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

CATALOG_URL = "https://raw.githubusercontent.com/gaborjager/odg-atlas/refs/heads/main/output/maps.json"
FOLDER = Path(__file__).resolve().parent
ATLAS = Path(os.environ.get("ODG_ATLAS") or FOLDER.parent.parent / "odg-atlas")

NS_DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
NS_SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"


def download(url):
    with urllib.request.urlopen(url, timeout=60) as r:
        return r.read()


def read(url):
    """File of the atlas: local copy if any (same path as on the main branch),
    else downloaded."""
    local = ATLAS / url.split("/main/", 1)[1]
    return local.read_bytes() if local.exists() else download(url)


def zones(odg):
    """Map codes -> name (alternative text without the code), in order."""
    with zipfile.ZipFile(odg) as z:
        root = ET.fromstring(z.read("content.xml"))
    result = {}
    for el in root.iter():
        code = el.get(f"{{{NS_DRAW}}}name")
        desc = el.find(f"{{{NS_SVG}}}desc")
        if not code or desc is None or code in result:
            continue
        if el.findtext(f"{{{NS_SVG}}}title") == "optional-zone":   # not a zone to test
            continue
        text = (desc.text or "").strip()
        # "22 - Côtes-d'Armor" -> "Côtes-d'Armor"
        name = text.split(" - ", 1)[1] if text.startswith(code + " - ") else text
        result[code] = name
    return result


def text_cell(s):
    return f'<table:table-cell office:value-type="string"><text:p>{escape(s)}</text:p></table:table-cell>'


def number_cell(v):
    return f'<table:table-cell office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:table-cell>'


def write_ods(path, sheet, rows):
    """rows: [(code, value, name)], preceded by a header row."""
    body = ["<table:table-row>" + text_cell("Code") + text_cell("Value (test)")
            + text_cell("Name") + "</table:table-row>"]
    for code, value, name in rows:
        body.append("<table:table-row>" + text_cell(code) + number_cell(value)
                    + text_cell(name) + "</table:table-row>")
    content = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<office:document-content'
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
        ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"'
        ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
        ' office:version="1.3"><office:body><office:spreadsheet>'
        f'<table:table table:name="{escape(sheet)}">'
        '<table:table-column table:number-columns-repeated="3"/>'
        + "".join(body)
        + "</table:table></office:spreadsheet></office:body></office:document-content>"
    )
    manifest = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
        '<manifest:file-entry manifest:full-path="/" manifest:version="1.3"'
        ' manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        "</manifest:manifest>"
    )
    with zipfile.ZipFile(path, "w") as z:
        # mimetype first and uncompressed (ODF requirement)
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet", zipfile.ZIP_STORED)
        z.writestr("content.xml", content, zipfile.ZIP_DEFLATED)
        z.writestr("META-INF/manifest.xml", manifest, zipfile.ZIP_DEFLATED)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv
    catalog = {c["id"]: c for c in json.loads(read(CATALOG_URL))["maps"]}
    ids = args or list(catalog)
    for id_ in ids:
        if id_ not in catalog:
            print(f"{id_}: not in the catalog ({', '.join(catalog)})")
            continue
        path = FOLDER / f"test-{id_}.ods"
        if path.exists() and not force:
            print(f"{path.name} already exists (--force to rewrite it)")
            continue
        entry = catalog[id_]
        print(f"{entry['filename']}: reading…", flush=True)
        codes = zones(io.BytesIO(read(entry["url"])))
        rng = random.Random(id_)            # reproducible values per map
        rows = [(code, rng.randint(0, 1000), name) for code, name in sorted(codes.items())]
        write_ods(path, id_, rows)
        print(f"{path.name}: {len(rows)} zones")


if __name__ == "__main__":
    main()
