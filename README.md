# CalcMap — cartes thématiques pour LibreOffice Calc

**CalcMap** est une extension LibreOffice qui colorie un fond de carte `.odg` à partir de deux colonnes
d'un classeur Calc : **code | valeur**. Elle utilise les fonds de carte du projet
[france-odg](https://github.com/gaborjager/france-odg), où chaque zone porte son
code INSEE dans son nom.

## Utilisation

1. Dans Calc, sélectionner les données : code en 1re colonne, valeur en 2e,
   libellé (facultatif) en 3e. Il suffit de sélectionner les deux premières : la
   colonne suivante est toujours lue comme libellé. Les autres colonnes sont
   ignorées ; on peut sélectionner des colonnes entières.
2. Cliquer sur l'icône **globe** de la barre d'outils (ou **Outils → CalcMap : colorier une carte…**).
3. Choisir le fond de carte dans la liste déroulante (régions, départements),
   ou un autre fichier `.odg` avec **Autre fichier…**.
4. Choisir la palette de couleurs (un aperçu s'affiche à côté de la liste).
5. Laisser cochée **Insérer la carte dans la feuille Calc** pour placer la carte
   dans la feuille, ou la décocher pour l'ouvrir dans Draw.
6. Cocher **Afficher les valeurs sur la carte** et/ou **Afficher les libellés
   (3e colonne)** pour écrire la valeur et/ou le libellé de chaque zone sur la
   carte, et choisir leur taille (voir ci-dessous).

> Les cartes des cantons et des communes du catalogue france-odg ne sont pas
> proposées pour l'instant : elles contiennent trop de zones (constante
> `CARTES_IGNOREES` du code).

La carte s'ouvre dans Draw, coloriée en classes d'intervalles égaux (autant de
classes que la palette compte de couleurs : 5, 7 ou 10), ou en dégradé continu
pour les palettes « dégradé continu ». Les zones sans
donnée restent en gris. La dernière carte et la dernière palette choisies sont
proposées par défaut à l'exécution suivante.

## Carte insérée dans la feuille Calc

Case cochée (par défaut), la carte coloriée est placée dans la feuille, à droite
du bloc de données sélectionné, après une colonne vide, sur 16 cm de large.
Elle est ancrée à sa cellule.

- **Jusqu'à 5 000 zones** (régions, départements) : les zones sont
  recréées dans la feuille et groupées. Chaque zone garde sa couleur et son texte
  alternatif (`22 - Côtes-d'Armor - Valeur : 666`). Pour la consulter : double-clic
  sur la carte pour entrer dans le groupe, puis clic droit sur une zone →
  *Texte alternatif*.
- **Au-delà** (fond de carte choisi avec *Autre fichier…*) : la carte est insérée
  comme **image** PNG. Les valeurs par zone ne sont alors pas consultables sur la
  carte ; décocher la case pour ouvrir la carte dans Draw si besoin.

Le choix de la case est mémorisé.

### Valeurs et libellés sur la carte

Deux cases du dialogue écrivent du texte au centre de chaque zone, dans la carte
insérée comme dans Draw : **Afficher les libellés (3e colonne)** (libellé sur la
1re ligne) et **Afficher les valeurs sur la carte** (valeur en dessous).

- La valeur est au **format de nombre de la colonne Calc** (décimales,
  pourcentage, séparateur de milliers…) : formater la colonne des valeurs avant
  de lancer la commande. Le libellé est repris tel qu'il s'affiche.
- Le texte est noir sur les couleurs claires, blanc sur les couleurs foncées.
- Chacun a sa **taille**, à côté de sa case (10 pt par défaut, de 2 à 72) :
  flèches pour +1 / −1, ou saisie directe. La taille vaut pour la carte insérée
  dans Calc (16 cm de large) ; dans Draw, elle est adaptée pour garder le même
  aspect.
- Une zone sans valeur n'affiche que son libellé.

Valeurs et libellés suivent la mise à jour automatique (un libellé modifié dans
la 3e colonne est aussi repris). Au-delà de 5 000 zones, aucun texte n'est
affiché (trop long à créer et illisible). Un texte peut être en partie masqué par
une zone voisine, et les petites zones (Paris et sa petite couronne) sont peu
lisibles. Les choix sont mémorisés ; les deux cases sont décochées par défaut.

### Configuration d'une carte insérée

**Clic droit sur la carte → Configuration de la carte…** rouvre le dialogue pour
cette carte : palette, valeurs, libellés et leurs tailles. La carte est recoloriée
dès la validation, et ces choix servent aussi de point de départ pour la carte
suivante. Le fond de carte ne peut pas être changé (insérer une nouvelle carte).
Le menu fonctionne aussi sur une zone, après un double-clic pour entrer dans la
carte.

### Mise à jour automatique

Une carte insérée dans la feuille reste **liée à ses données** : dès qu'une
valeur change (saisie, collage, résultat de formule recalculé), la carte est
recoloriée, avec la même palette (et les mêmes valeurs affichées ou non). Pour une
carte insérée en image, l'image est régénérée.

- Le lien est enregistré dans le classeur : une plage nommée `CalcMap_1`
  (`_2`, … pour d'autres cartes) sur les deux colonnes sélectionnées, et une
  propriété du document du même nom (*Fichier → Propriétés → Propriétés
  personnalisées*). La plage nommée suit les insertions de lignes ; sélectionner
  des colonnes entières (A:B) prend aussi en compte les lignes ajoutées en bas.
- À l'ouverture du classeur, l'extension reconnecte ses cartes (événement
  « Ouvrir le document », déclaré dans `Events.xcu`). Sans l'extension installée,
  le classeur s'ouvre normalement ; la carte reste simplement figée.
- Le recoloriage n'entre pas dans l'historique : *Annuler* annule la saisie, et
  l'insertion de la carte est une seule étape d'annulation.
- Pour rompre le lien : supprimer la carte de la feuille (le lien est retiré à la
  modification suivante), ou supprimer la plage nommée `CalcMap_N`.

> L'extension occupe l'événement global « Ouvrir le document »
> (*Outils → Personnaliser → Événements*, « Enregistrer dans : LibreOffice »). Si
> une autre macro y était déjà associée, elle est remplacée.

## Palettes de couleurs

Palettes fournies (d'après ColorBrewer et viridis) :

| Type | Palettes |
|---|---|
| 5 classes | bleus, verts, rouges, oranges, violets, gris, jaune → rouge, viridis (lisible par les daltoniens), divergentes rouge ↔ bleu et brun ↔ vert |
| 7 classes | bleus |
| 10 classes | bleus, verts, jaune → rouge, viridis, divergentes rouge ↔ bleu et spectrale |
| Dégradé continu | bleus, jaune → rouge, viridis, gris, divergente rouge ↔ bleu |

En **dégradé continu**, il n'y a pas de classes : chaque zone reçoit une couleur
interpolée selon sa valeur exacte, entre le minimum (première couleur) et le
maximum (dernière couleur). Les couleurs intermédiaires de la palette sont des
étapes du dégradé (utile pour une palette divergente : bas → milieu → haut).

**Ajouter ses palettes**, sans reconstruire l'extension : créer un fichier
`palettes.json` dans le dossier du cache (voir ci-dessous). Ses palettes s'ajoutent
à la liste ; une palette de même `id` qu'une palette fournie la remplace.

```json
{
  "palettes": [
    { "id": "ma-palette",
      "name": { "fr": "Ma palette", "en": "My palette" },
      "colors": ["#fff5eb", "#fdae6b", "#e6550d", "#7f2704"],
      "nodata": "#ffffff" },
    { "id": "mon-degrade",
      "name": "Mon dégradé",
      "continuous": true,
      "colors": ["#ffffff", "#006d2c"] }
  ]
}
```

- `colors` : de la plus petite à la plus grande valeur, 2 couleurs au moins,
  au format `#rrggbb`. Leur nombre donne le nombre de classes.
- `name` : texte simple, ou un texte par langue (repli sur `en`, puis sur `id`).
- `nodata` (facultatif) : couleur des zones sans donnée, gris clair par défaut.
- `continuous` (facultatif) : `true` pour un dégradé continu ; les couleurs sont
  alors les étapes du dégradé et non des classes.

Une palette mal formée est ignorée et signalée dans le journal. Les palettes de
l'extension sont dans `src/palettes.json`.

## Langues

L'interface suit la langue de LibreOffice : français et anglais sont fournis, et
l'anglais sert de repli pour tout texte non traduit. Les textes sont dans
`src/i18n/messages_<langue>.properties` (UTF-8, `clé=texte`, `\n` pour un retour à
la ligne, `%1` pour une valeur insérée). **Ajouter une langue** : copier
`messages_en.properties` en `messages_<code>.properties` (ex. `messages_de.properties`),
traduire, puis reconstruire l'extension ; `build.py` inclut tout le dossier `i18n/`.
Le catalogue france-odg peut aussi fournir des champs traduits (`title_en`,
`description_en`…), utilisés quand ils existent. Le journal reste en français.

## Fonds de carte et cache

La liste des cartes vient du catalogue
[`maps.json`](https://raw.githubusercontent.com/gaborjager/france-odg/refs/heads/main/output/maps.json)
de france-odg, lu à chaque exécution. Les cartes sont téléchargées à la demande
dans le profil LibreOffice :

- Windows : `%APPDATA%\LibreOffice\4\user\calcmap\`
- Linux : `~/.config/libreoffice/4/user/calcmap/`

Avant chaque utilisation, l'empreinte SHA-256 du fichier en cache est comparée à
celle du catalogue : la carte n'est retéléchargée que si elle a changé. Un
téléchargement dont l'empreinte ne correspond pas est rejeté. La carte s'ouvre
comme un nouveau document (modèle) : l'enregistrer ne modifie jamais le cache.

**Sans réseau**, la liste ne propose que les cartes déjà téléchargées (ainsi que
tout autre `.odg` déposé dans ce dossier). Le réseau passe par LibreOffice :
un proxy se règle dans **Outils → Options → Internet → Proxy**.

## Comportement face aux données imparfaites

La macro n'affiche jamais d'erreur Basic. Une donnée inexploitable laisse la zone
en couleur par défaut (gris) et est notée dans le journal.

| Cas | Traitement |
|---|---|
| Ligne vide, code vide | ignorée |
| En-tête (même répété au milieu) | ignoré (valeur non numérique) |
| Valeur vide, texte (`n.d.`), cellule en erreur (`#DIV/0!`) | couleur par défaut |
| Nombre saisi en texte (`1 234,5`) | accepté |
| Code sans zéro initial (`1` pour `01`) | retrouvé |
| Code en minuscules ou avec espaces (` 2a `) | accepté |
| Code en double | dernière valeur retenue |
| Code absent de la carte | ignoré |

## Texte alternatif

La valeur de chaque zone est ajoutée à son **texte alternatif** (clic droit sur
une zone → *Texte alternatif*, champ *Texte alternatif*), à la suite du texte
fourni par la carte :

- `22 - Côtes-d'Armor` devient `22 - Côtes-d'Armor - Valeur : 666`
- sans donnée : `22 - Côtes-d'Armor - Valeur : (aucune)`

Recolorier une carte déjà coloriée remplace la valeur au lieu de l'ajouter une
seconde fois. Désactivable avec `ECRIRE_TEXTE_ALT = False` dans le code.

## Journal (débogage)

Chaque exécution écrase le fichier `calcmap.log` du profil LibreOffice :

- Windows : `%APPDATA%\LibreOffice\4\user\calcmap.log`
- Linux : `~/.config/libreoffice/4/user/calcmap.log`

Il liste la langue, les palettes, la plage lue, le catalogue et le cache (empreintes, téléchargements), chaque ligne écartée avec sa raison, les codes inconnus de
la carte et le bilan. Avec `JOURNAL_DETAILLE = True`, il trace aussi chaque zone.

## Installation

Télécharger [`dist/CalcMap.oxt`](https://github.com/gaborjager/calcmap/raw/main/dist/CalcMap.oxt)
(dernière version construite), puis :

- **Un utilisateur** : double-clic sur `CalcMap.oxt`, ou
  **Outils → Gestionnaire des extensions → Ajouter**, puis redémarrer LibreOffice.
- **Tous les utilisateurs d'un poste** (administrateur, LibreOffice fermé) :
  ```
  "C:\Program Files\LibreOffice\program\unopkg.exe" add --shared CalcMap.oxt
  ```

Le bouton s'ajoute à la barre d'outils *Standard* de Calc, après *Insérer un graphique*.
Avec l'interface en onglets, utiliser **Outils → CalcMap : colorier une carte…**.

### Mise à jour depuis ColorierCarte

CalcMap est le nouveau nom de l'extension *ColorierCarte*. C'est
une nouvelle extension pour LibreOffice : **désinstaller ColorierCarte** dans le
Gestionnaire des extensions avant d'installer CalcMap, sinon les deux boutons
coexistent. Rien n'est perdu :

- au premier lancement, le cache `coloriercarte/` (cartes téléchargées, palettes
  personnelles `palettes.json`, derniers choix) est recopié dans `calcmap/` ;
- les classeurs dont la carte est liée par ColorierCarte (plages `ColorierCarte_N`)
  continuent de se mettre à jour ; les nouvelles cartes utilisent `CalcMap_N`.

## Construire l'extension

```bash
python3 build.py
```

Produit `dist/CalcMap.oxt`. La version est lue dans la constante `VERSION`
du code Basic et reportée dans `description.xml`.

## Fichiers de test

Un classeur par fond de carte du catalogue (ceux des cantons et des communes
sont conservés, bien que ces cartes ne soient pas proposées pour l'instant) : `tests/test-<carte>.ods`, avec les
colonnes *Code | Valeur (test) | Nom* et une ligne par zone. Sélectionner les
colonnes A:B puis lancer la commande avec la carte correspondante : toutes les
zones doivent être coloriées, sans code inconnu dans le journal.

Pour les régénérer depuis les cartes du catalogue (valeurs aléatoires mais
reproductibles) :

```bash
python3 tests/generer_tests.py                        # fichiers manquants
python3 tests/generer_tests.py departements --force   # réécrire une carte
```

## Organisation

```
calcmap/
├── build.py                    construit dist/CalcMap.oxt
├── dist/CalcMap.oxt            extension prête à installer
├── src/
│   ├── basic/CalcMap.bas       code source de la macro (à modifier ici)
│   ├── CalcMap/                bibliothèque Basic (le .xba est généré par build.py)
│   ├── palettes.json           palettes de couleurs fournies
│   ├── i18n/                   textes traduits (messages_fr, messages_en…)
│   ├── description/            description de l'extension par langue
│   ├── Addons.xcu              bouton, entrée de menu, icône
│   ├── Events.xcu              reconnexion des cartes liées à l'ouverture
│   ├── description.xml         identifiant et version de l'extension
│   ├── META-INF/manifest.xml
│   └── icons/                  icône globe 16 et 26 px
└── tests/
    ├── generer_tests.py        (re)génère les classeurs test-<carte>.ods
    ├── test-regions.ods        13 régions        ┐
    ├── test-departements.ods   96 départements   │ Code | Valeur (test) | Nom,
    ├── test-cantons.ods        2 294 cantons     │ une ligne par zone de la carte
    ├── test-communes.ods       35 191 communes   ┘
    ├── test-cas-limites.ods    en-têtes, vides, textes, erreurs, doublons…
    └── apercu-test-departements.png
```
