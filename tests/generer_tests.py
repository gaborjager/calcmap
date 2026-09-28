#!/usr/bin/env python3
"""Génère un classeur de test (.ods) par fond de carte du catalogue france-odg.

Usage :
    python3 tests/generer_tests.py                 # régions, cantons, communes
    python3 tests/generer_tests.py departements    # cartes choisies (id du catalogue)
    python3 tests/generer_tests.py --force         # réécrit aussi les fichiers existants

Chaque classeur contient : Code | Valeur (test) | Nom, une ligne par zone de la
carte. Les codes et les noms sont lus dans la carte elle-même (draw:name et
texte alternatif svg:desc) ; les valeurs sont aléatoires mais reproductibles.
Les codes sont écrits en texte pour conserver les zéros initiaux (« 01 »).
"""
import io
import json
import random
import sys
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

URL_CATALOGUE = "https://raw.githubusercontent.com/gaborjager/france-odg/refs/heads/main/output/maps.json"
DOSSIER = Path(__file__).resolve().parent
PAR_DEFAUT = ["regions", "departements", "cantons", "communes"]

NS_DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
NS_SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"


def telecharger(url):
    with urllib.request.urlopen(url, timeout=60) as r:
        return r.read()


def zones(odg):
    """Codes de la carte -> nom (texte alternatif sans le code), dans l'ordre."""
    with zipfile.ZipFile(odg) as z:
        racine = ET.fromstring(z.read("content.xml"))
    resultat = {}
    for el in racine.iter():
        code = el.get(f"{{{NS_DRAW}}}name")
        desc = el.find(f"{{{NS_SVG}}}desc")
        if not code or desc is None or code in resultat:
            continue
        texte = (desc.text or "").strip()
        # « 22 - Côtes-d'Armor » -> « Côtes-d'Armor »
        nom = texte.split(" - ", 1)[1] if texte.startswith(code + " - ") else texte
        resultat[code] = nom
    return resultat


def cellule_texte(s):
    return f'<table:table-cell office:value-type="string"><text:p>{escape(s)}</text:p></table:table-cell>'


def cellule_nombre(v):
    return f'<table:table-cell office:value-type="float" office:value="{v}"><text:p>{v}</text:p></table:table-cell>'


def ecrire_ods(chemin, feuille, lignes):
    """lignes : [(code, valeur, nom)], précédées d'une ligne d'en-tête."""
    corps = ["<table:table-row>" + cellule_texte("Code") + cellule_texte("Valeur (test)")
             + cellule_texte("Nom") + "</table:table-row>"]
    for code, valeur, nom in lignes:
        corps.append("<table:table-row>" + cellule_texte(code) + cellule_nombre(valeur)
                     + cellule_texte(nom) + "</table:table-row>")
    contenu = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<office:document-content'
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
        ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"'
        ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
        ' office:version="1.3"><office:body><office:spreadsheet>'
        f'<table:table table:name="{escape(feuille)}">'
        '<table:table-column table:number-columns-repeated="3"/>'
        + "".join(corps)
        + "</table:table></office:spreadsheet></office:body></office:document-content>"
    )
    manifeste = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
        '<manifest:file-entry manifest:full-path="/" manifest:version="1.3"'
        ' manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
        '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        "</manifest:manifest>"
    )
    with zipfile.ZipFile(chemin, "w") as z:
        # mimetype en premier et non compressé (exigence ODF)
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet", zipfile.ZIP_STORED)
        z.writestr("content.xml", contenu, zipfile.ZIP_DEFLATED)
        z.writestr("META-INF/manifest.xml", manifeste, zipfile.ZIP_DEFLATED)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv
    catalogue = {c["id"]: c for c in json.loads(telecharger(URL_CATALOGUE))["maps"]}
    ids = args or PAR_DEFAUT
    for id_ in ids:
        if id_ not in catalogue:
            print(f"{id_} : absent du catalogue ({', '.join(catalogue)})")
            continue
        chemin = DOSSIER / f"test-{id_}.ods"
        if chemin.exists() and not force:
            print(f"{chemin.name} existe déjà (--force pour le réécrire)")
            continue
        carte = catalogue[id_]
        print(f"{carte['filename']} : téléchargement…", flush=True)
        codes = zones(io.BytesIO(telecharger(carte["url"])))
        alea = random.Random(id_)            # valeurs reproductibles par carte
        lignes = [(code, alea.randint(0, 1000), nom) for code, nom in sorted(codes.items())]
        ecrire_ods(chemin, id_, lignes)
        print(f"{chemin.name} : {len(lignes)} zones")


if __name__ == "__main__":
    main()
