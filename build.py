#!/usr/bin/env python3
"""Construit l'extension LibreOffice CalcMap.oxt à partir de src/.

Usage : python3 build.py
Résultat : dist/CalcMap.oxt

Le code Basic est maintenu dans src/basic/CalcMap.bas (lisible, diffable) ;
ce script le convertit au format .xba attendu par LibreOffice.
"""
import html
import re
import zipfile
from pathlib import Path

RACINE = Path(__file__).resolve().parent
SRC = RACINE / "src"
DIST = RACINE / "dist"

EN_TETE_XBA = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
    '<script:module xmlns:script="http://openoffice.org/2000/script" '
    'script:name="CalcMap" script:language="StarBasic">'
)

FICHIERS = [
    "META-INF/manifest.xml",
    "description.xml",
    "Addons.xcu",
    "Events.xcu",
    "CalcMap/script.xlb",
    "CalcMap/dialog.xlb",
    "icons/carte_16.png",
    "icons/carte_26.png",
    "icons/nopreview.svg",
    "palettes.json",
    "CalcMap.xcs",
]
# Dossiers inclus en entier : ajouter une langue = ajouter un fichier
DOSSIERS = ["i18n", "description"]


def main():
    basic = (SRC / "basic" / "CalcMap.bas").read_text(encoding="utf-8")

    # La version de description.xml doit suivre celle du code Basic
    m = re.search(r'Const VERSION = "([\d.]+)"', basic)
    version = m.group(1) if m else "0.0.0"
    desc = (SRC / "description.xml").read_text(encoding="utf-8")
    desc = re.sub(r'<version value="[^"]*"/>', f'<version value="{version}"/>', desc)
    (SRC / "description.xml").write_text(desc, encoding="utf-8")

    xba = EN_TETE_XBA + html.escape(basic, quote=False) + "</script:module>\n"
    (SRC / "CalcMap" / "CalcMap.xba").write_text(xba, encoding="utf-8")

    DIST.mkdir(exist_ok=True)
    sortie = DIST / "CalcMap.oxt"
    with zipfile.ZipFile(sortie, "w", zipfile.ZIP_DEFLATED) as z:
        extras = [p.relative_to(SRC).as_posix() for d in DOSSIERS
                  for p in sorted((SRC / d).iterdir()) if p.is_file()]
        for f in FICHIERS + extras + ["CalcMap/CalcMap.xba"]:
            z.write(SRC / f, f)
        # Catalogue agrégé du projet : secours si celui configuré est injoignable
        z.write(RACINE / "catalogue" / "catalogue.json", "catalogue.json")
    print(f"{sortie.relative_to(RACINE)} (version {version})")


if __name__ == "__main__":
    main()
