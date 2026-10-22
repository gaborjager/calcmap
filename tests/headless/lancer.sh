#!/bin/bash
# Lance une macro de test de CalcMap dans LibreOffice sans écran.
#
# Usage : tests/headless/lancer.sh <macro> [classeur.ods]
#   macro    : Sub de tests_calcmap.bas (ex. TestEditeurComplet)
#   classeur : par défaut tests/test-departements.ods (copié, jamais modifié)
#
# Construit une copie de l'extension avec les macros de test ajoutées,
# l'installe dans un profil LibreOffice jetable, lance la macro et affiche
# les lignes « TEST » du journal. Image de la carte : $TRAVAIL/carte.png.
set -e
ICI="$(cd "$(dirname "$0")" && pwd)"
PROJET="$(cd "$ICI/../.." && pwd)"
MACRO="${1:?macro de test attendue}"
CLASSEUR="${2:-$PROJET/tests/test-departements.ods}"
TRAVAIL="${TMPDIR:-/tmp}/calcmap-tests"
PORT=$((2200 + RANDOM % 500))

rm -rf "$TRAVAIL/copie" "$TRAVAIL/profil"
mkdir -p "$TRAVAIL"
cp -r "$PROJET" "$TRAVAIL/copie"
rm -rf "$TRAVAIL/copie/.git"
cat "$ICI/tests_calcmap.bas" >> "$TRAVAIL/copie/src/basic/CalcMap.bas"
(cd "$TRAVAIL/copie" && python3 build.py > /dev/null)
unopkg add -f "-env:UserInstallation=file://$TRAVAIL/profil" "$TRAVAIL/copie/dist/CalcMap.oxt" > /dev/null 2>&1
cp "$CLASSEUR" "$TRAVAIL/classeur.ods"
rm -f "$TRAVAIL"/.~lock*
timeout 300 python3 "$ICI/pilote.py" "$TRAVAIL/profil" "$TRAVAIL/classeur.ods" "$MACRO" "$TRAVAIL/carte.png" "$PORT" \
	> /dev/null 2>&1 || true
grep -E "TEST|ERREUR" "$TRAVAIL/profil/user/calcmap.log" | cut -c22-
