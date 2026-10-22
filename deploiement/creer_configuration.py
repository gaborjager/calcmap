#!/usr/bin/env python3
"""Crée une extension de configuration de CalcMap pour un déploiement.

L'extension produite (CalcMap-configuration.oxt) ne contient que les
paramètres de CalcMap : catalogue agrégé et nombre maximal de zones. Un
administrateur l'installe pour tous les utilisateurs d'un poste, après
CalcMap :

    unopkg add --shared CalcMap-configuration.oxt

Usage :
    python3 deploiement/creer_configuration.py --catalogue URL [--zones-max N]
        [--verrouiller] [--sortie dist/CalcMap-configuration.oxt]

    --catalogue    URL du catalogue agrégé (https://… ou file:///…)
    --zones-max    cartes proposées : au plus ce nombre de zones
    --verrouiller  les utilisateurs ne peuvent pas modifier ces paramètres
"""
import argparse
import html
import zipfile
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent

DESCRIPTION = """<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/description/2006"
             xmlns:d="http://openoffice.org/extensions/description/2006"
             xmlns:xlink="http://www.w3.org/1999/xlink">
 <identifier value="fr.github.gaborjager.calcmap.configuration"/>
 <version value="1.0.0"/>
 <display-name>
  <name lang="fr">CalcMap - configuration ({nom})</name>
  <name lang="en">CalcMap - configuration ({nom})</name>
 </display-name>
</description>
"""

MANIFESTE = """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
 <manifest:file-entry manifest:full-path="CalcMapConfiguration.xcu"
   manifest:media-type="application/vnd.sun.star.configuration-data"/>
</manifest:manifest>
"""


def parametre(nom, type_xs, valeur, verrouiller):
    final = ' oor:finalized="true"' if verrouiller else ""
    return (f'  <prop oor:name="{nom}" oor:type="{type_xs}" oor:op="fuse"{final}>\n'
            f'   <value>{html.escape(str(valeur), quote=False)}</value>\n'
            f'  </prop>\n')


def main():
    p = argparse.ArgumentParser(description="Extension de configuration de CalcMap pour un déploiement")
    p.add_argument("--catalogue", help="URL du catalogue agrégé (https://… ou file:///…)")
    p.add_argument("--zones-max", type=int, help="nombre maximal de zones des cartes proposées")
    p.add_argument("--verrouiller", action="store_true", help="empêcher les utilisateurs de modifier ces paramètres")
    p.add_argument("--sortie", default=str(RACINE / "dist" / "CalcMap-configuration.oxt"))
    a = p.parse_args()
    if a.catalogue is None and a.zones_max is None:
        p.error("indiquer --catalogue et/ou --zones-max")
    if a.catalogue and not a.catalogue.lower().startswith(("https://", "file:///")):
        p.error("le catalogue doit être une URL https:// ou file:///")
    if a.zones_max is not None and a.zones_max <= 0:
        p.error("--zones-max doit être positif")

    props = ""
    if a.catalogue:
        props += parametre("Catalogue", "xs:string", a.catalogue, a.verrouiller)
    if a.zones_max is not None:
        props += parametre("ZonesMax", "xs:int", a.zones_max, a.verrouiller)
    xcu = ('<?xml version="1.0" encoding="UTF-8"?>\n'
           '<oor:component-data xmlns:oor="http://openoffice.org/2001/registry"\n'
           '    xmlns:xs="http://www.w3.org/2001/XMLSchema"\n'
           '    oor:name="CalcMap" oor:package="fr.github.gaborjager">\n'
           ' <node oor:name="Parametres">\n' + props + ' </node>\n</oor:component-data>\n')

    sortie = Path(a.sortie)
    sortie.parent.mkdir(parents=True, exist_ok=True)
    nom = "verrouillée" if a.verrouiller else "par défaut"
    with zipfile.ZipFile(sortie, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("description.xml", DESCRIPTION.format(nom=nom))
        z.writestr("META-INF/manifest.xml", MANIFESTE)
        z.writestr("CalcMapConfiguration.xcu", xcu)
    print(f"{sortie} : catalogue={a.catalogue or '(inchangé)'}, zones max={a.zones_max or '(inchangé)'}, "
          f"{'verrouillé' if a.verrouiller else 'modifiable par les utilisateurs'}")


if __name__ == "__main__":
    main()
