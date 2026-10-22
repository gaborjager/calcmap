# CalcMap — cartes thématiques pour LibreOffice Calc

**CalcMap** est une extension LibreOffice qui colorie un fond de carte `.odg` à partir de deux colonnes
d'un classeur Calc : **code | valeur**. Elle utilise les fonds de carte du projet
[france-odg](https://github.com/gaborjager/france-odg), où chaque zone porte son
code INSEE dans son nom.

## Utilisation

1. Dans Calc, sélectionner les données, en-têtes compris : par défaut code en
   1re colonne, valeur en 2e, libellé (facultatif) en 3e. Si deux colonnes
   seulement sont sélectionnées, la suivante est proposée comme libellé. On peut
   sélectionner des colonnes entières.
2. Cliquer sur l'icône **globe** de la barre d'outils (ou **Outils → CalcMap : colorier une carte…**).
3. Choisir le fond de carte dans la liste déroulante (régions, départements),
   ou un autre fichier `.odg` avec **Autre fichier…**.
4. Choisir la palette de couleurs (un aperçu s'affiche à côté de la liste), et
   au besoin personnaliser les classes avec **Classes…** (voir ci-dessous).
5. Vérifier les colonnes du **code**, des **valeurs** et des **libellés** : chaque
   liste propose les colonnes sélectionnées, avec leur en-tête et leurs premières
   valeurs (ex. « D — Taux : 12,5, 8,1, 9,7 »).
6. Laisser cochée **Insérer la carte dans la feuille Calc** pour placer la carte
   dans la feuille, ou la décocher pour l'ouvrir dans Draw.
7. Cocher **Afficher les valeurs** et/ou **Afficher les libellés**
   pour écrire la valeur et/ou le libellé de chaque zone sur la carte, et choisir
   leur taille, leur police et leur style (voir ci-dessous).
8. Cocher **Afficher la légende**, **avec les valeurs** (bornes de chaque classe)
   ou sans.

> Seules les cartes d'au plus **300 zones** sont proposées (paramètre `ZonesMax`,
> voir *Paramètres*) : avec le catalogue actuel, les régions et les départements
> (et à terme une carte des pays du monde). Les cantons (2 306 zones) et les
> communes (35 191) sont écartés.

Le dialogue est organisé en cadres : **Carte** (fond de carte, son aperçu,
insertion dans la feuille), **Code** (colonne), **Valeurs** (colonne, couleurs et classes,
affichage des valeurs), **Libellés** (colonne, affichage) et **Légende**.

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
insérée comme dans Draw : **Afficher les libellés** (libellé sur la 1re ligne) et
**Afficher les valeurs** (valeur en dessous).

- La valeur est au **format de nombre de la colonne Calc** (décimales,
  pourcentage, séparateur de milliers…) : formater la colonne des valeurs avant
  de lancer la commande. Le libellé est repris tel qu'il s'affiche.
- Le texte est noir sur les couleurs claires, blanc sur les couleurs foncées.
- À côté de chaque case, comme dans la barre de formatage : la **police**
  (Liberation Sans par défaut), la **taille** (10 pt par défaut, de 2 à 72),
  choisies dans la liste ou saisies, et les boutons **G** (gras), *I* (italique),
  S (souligné). La taille vaut pour la carte insérée dans Calc (16 cm de large) ;
  dans Draw, elle est adaptée pour garder le même aspect.
- Une zone sans valeur n'affiche que son libellé.
- **Contour blanc ou noir** (une case sous chaque ligne, cochée par défaut) :
  texte noir cerné de blanc, texte blanc cerné de noir, lisible sur toutes les
  couleurs ; réglable séparément pour les valeurs et pour les libellés. Le
  contour est fait de copies du texte décalées tout autour : une mise à jour
  complète prend alors 3 s environ pour les départements, au lieu de 0,6 s.

Valeurs et libellés suivent la mise à jour automatique (un libellé modifié est
aussi repris). Au-delà de 5 000 zones, aucun texte n'est
affiché (trop long à créer et illisible). Les textes sont des formes séparées,
placées au-dessus de toutes les zones : une zone voisine ne les masque jamais.
Dans les petites zones (Paris et sa petite couronne), ils peuvent se chevaucher. Les choix sont mémorisés ; les deux cases sont décochées par défaut.

### Légende

Case **Afficher la légende** cochée, une légende est placée à droite de la
carte : une case de couleur par classe (5 nuances pour un dégradé continu), et
une case « Sans donnée » si des zones n'ont pas de valeur.

- **avec les valeurs** : chaque case porte les bornes de sa classe
  (`20 – 411,2`), au format de la colonne des valeurs, et la légende a pour titre
  l'en-tête de cette colonne ; sinon, seules les couleurs sont affichées.
- Le texte reprend la police et la taille des valeurs.
- Dans Calc, la légende fait partie de la carte : elle se déplace avec elle et
  se met à jour avec elle.

### Configuration d'une carte insérée

**Clic droit sur la carte → Configuration de la carte…**, ou l'**icône globe**
quand la carte (ou une de ses zones) est sélectionnée, rouvre le dialogue pour
cette carte : palette, colonnes, valeurs, libellés, légende. La carte est
recoloriée dès la validation, et ces choix (sauf les colonnes, propres à chaque
tableau) servent aussi de point de départ pour la carte suivante. Le fond de carte ne peut pas être changé (insérer une nouvelle carte).
Le menu fonctionne aussi sur une zone, après un double-clic pour entrer dans la
carte.

### Mise à jour automatique

Une carte insérée dans la feuille reste **liée à ses données** : dès qu'une
valeur change (saisie, collage, résultat de formule recalculé), la carte est
mise à jour, avec les mêmes réglages.

- En général, **seules les zones modifiées** sont refaites (couleur, texte
  alternatif, valeur et libellé affichés) : 0,3 s environ pour un département,
  contour compris.
- **Toute la carte** est refaite quand la modification touche toute la carte :
  nouveau minimum ou maximum (les classes linéaires bougent), zone qui gagne ou
  perd sa valeur (ligne « Sans donnée » de la légende), en-tête de la colonne
  des valeurs (titre de la légende), plus de 20 zones changées à la fois, ou
  nouveaux réglages (*Configuration de la carte…*). Pour une carte insérée en
  image, l'image est régénérée.

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

### Classes personnalisées

Par défaut, les valeurs sont réparties en classes d'intervalles égaux, une par
couleur de la palette. Le bouton **Classes…** (à droite de la palette) ouvre
l'éditeur de classes, sur les données de la colonne des valeurs choisie :

- **Nombre de classes** (flèches ou saisie, de 2 à 20) : les classes sont
  réparties à intervalles égaux, couleurs de la palette étirées en dégradé ;
- une **barre de couleurs** à l'échelle des valeurs : **glisser** un séparateur
  vers la gauche ou la droite déplace la borne entre deux classes (elle reste
  entre ses voisines) ; la valeur de chaque borne est affichée dessous ;
  **cliquer** sur une couleur ouvre le sélecteur de couleur ;
- chaque case affiche son **nombre de zones**, et sous la barre un trait par
  valeur montre où se concentrent les données ; la liste de droite donne toutes
  les valeurs triées, avec le libellé de chaque zone ;
- **Borne sélectionnée** : la valeur exacte du dernier séparateur touché, à
  saisir au clavier ;
- **Répartition linéaire** remet des intervalles égaux, en gardant le nombre de
  classes et les couleurs.

Les classes couvrent toujours toutes les valeurs (du minimum au maximum). Elles
sont gardées avec la carte (mise à jour automatique, légende, *Configuration de
la carte…*), mais pas pour la carte suivante. Choisir une autre palette remplace
les classes personnalisées.

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

### Catalogue

La liste des cartes vient d'un **catalogue agrégé**, lu à chaque exécution : par
défaut [`catalogue/catalogue.json`](catalogue/catalogue.json) de ce projet. Il liste
des **sources**, chacune un catalogue de cartes au format de
[france-odg](https://github.com/gaborjager/france-odg) (`maps.json`) ; pour
l'instant une seule, france-odg.

```json
{
  "version": 1,
  "title": "Catalogue CalcMap",
  "sources": [
    { "id": "france-odg", "title": "Cartes de France (france-odg)",
      "url": "https://raw.githubusercontent.com/gaborjager/france-odg/main/output/maps.json" }
  ]
}
```

- Une source par catalogue de cartes ; `id` (lettres, chiffres, `-`, `_`) nomme sa
  copie en cache. Une carte dont l'identifiant est déjà fourni par une source
  précédente est ignorée.
- Le paramètre `Catalogue` peut aussi désigner directement un catalogue de cartes
  (`maps.json`).
- URL acceptées (catalogue, sources, cartes, aperçus) : `https://`, ou `file:///`
  pour un partage réseau.
- Champs utilisés d'une carte : `id`, `title`, `description`, `url`,
  `filename`, `size`, `sha256`, et s'ils existent `objects` (nombre de zones,
  pour le filtre) et `preview` / `preview_sha256` (aperçu SVG, PNG ou JPEG).

### Paramètres

Deux paramètres, dans la configuration de LibreOffice
(`/fr.github.gaborjager.CalcMap/Parametres`) :

| Paramètre | Par défaut | Rôle |
|---|---|---|
| `Catalogue` | `https://raw.githubusercontent.com/gaborjager/calcmap/main/catalogue/catalogue.json` | catalogue agrégé (ou catalogue de cartes) |
| `ZonesMax` | `300` | cartes proposées : au plus ce nombre de zones |

Ils se modifient dans *Outils → Options → LibreOffice → Avancé → Ouvrir la
configuration avancée* : chercher `ZonesMax` ou `gaborjager` (chercher
« calcmap » trouve aussi de nombreux chemins de fichiers), puis double-cliquer sur
la valeur. Ils sont pris en compte à la prochaine exécution de CalcMap.

Une entreprise peut fournir son propre catalogue agrégé (ses sources : france-odg,
ses cartes internes…). Pour changer de sources ensuite, il suffit de republier ce
`catalogue.json` à la même adresse.

### Déploiement par stratégie de groupe (Windows)

LibreOffice lit aussi sa configuration dans le registre Windows, sous
`SOFTWARE\Policies\LibreOffice` : une stratégie de groupe peut donc imposer les
paramètres de CalcMap comme ceux de LibreOffice (*Préférences → Paramètres
Windows → Registre*, ou un modèle ADMX). Une clé par paramètre, sous
`HKEY_LOCAL_MACHINE` (tous les utilisateurs du poste) :

| Clé | Valeur `Value` (REG_SZ) | Valeur `Final` (REG_DWORD) |
|---|---|---|
| `SOFTWARE\Policies\LibreOffice\fr.github.gaborjager.CalcMap\Parametres\Catalogue` | `https://intranet.exemple.fr/cartes/catalogue.json` | `1` : l'utilisateur ne peut plus la modifier |
| `SOFTWARE\Policies\LibreOffice\fr.github.gaborjager.CalcMap\Parametres\ZonesMax` | `300` | `1` |

Sans `Final`, la valeur n'est qu'une valeur par défaut, que l'utilisateur peut
changer. CalcMap doit être installé (de préférence pour tous les utilisateurs :
`unopkg add --shared CalcMap.oxt`) ; les valeurs sont lues au démarrage de
LibreOffice. Pour vérifier sur un poste : dans la configuration avancée, la
valeur imposée apparaît, et n'est plus modifiable si elle est verrouillée.

> Mécanisme standard de LibreOffice, **non testé avec CalcMap** (pas de poste
> Windows ici) : à valider sur un poste de test avant le déploiement.

### Aperçu

La carte choisie dans la liste est montrée à droite (aperçu `preview` du
catalogue, téléchargé dans le cache et vérifié par son empreinte). Sans aperçu,
ou s'il est illisible, une image neutre est affichée.

### Cache

Le catalogue, ses sources, les cartes et les aperçus sont gardés dans le profil
LibreOffice :

- Windows : `%APPDATA%\LibreOffice\4\user\calcmap\`
- Linux : `~/.config/libreoffice/4/user/calcmap/`

Avant chaque utilisation, l'empreinte SHA-256 du fichier en cache est comparée à
celle du catalogue : la carte n'est retéléchargée que si elle a changé. Un
téléchargement dont l'empreinte ne correspond pas est rejeté. La carte s'ouvre
comme un nouveau document (modèle) : l'enregistrer ne modifie jamais le cache.

**Sans réseau**, les copies en cache du catalogue et des sources sont utilisées,
et la liste ne propose que les cartes déjà téléchargées (ainsi que tout autre
`.odg` déposé dans ce dossier). Au tout premier lancement sans réseau, le
catalogue fourni avec l'extension sert de secours. Le réseau passe par
LibreOffice : un proxy se règle dans **Outils → Options → Internet → Proxy**.

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
servent si le paramètre `ZonesMax` est relevé) : `tests/test-<carte>.ods`, avec les
colonnes *Code | Valeur (test) | Nom* et une ligne par zone. Sélectionner les
colonnes A:B puis lancer la commande avec la carte correspondante : toutes les
zones doivent être coloriées, sans code inconnu dans le journal.

Pour les régénérer depuis les cartes du catalogue (valeurs aléatoires mais
reproductibles) :

```bash
python3 tests/generer_tests.py                        # fichiers manquants
python3 tests/generer_tests.py departements --force   # réécrire une carte
```

### Tests automatiques (sans écran)

`tests/headless/lancer.sh <macro>` lance une macro de test de
`tests/headless/tests_calcmap.bas` dans LibreOffice sans écran, sur une copie de
l'extension et un profil jetable, et affiche les lignes `TEST` du journal :

```bash
tests/headless/lancer.sh TestTextes            # textes des zones, légende
tests/headless/lancer.sh TestEditeurComplet    # éditeur de classes (souris, flèches, saisie)
tests/headless/lancer.sh TestCatalogue         # catalogue agrégé, filtre, aperçus (fichiers locaux)
tests/headless/lancer.sh TestIconeSurCarte     # icône sur une carte sélectionnée : configuration
tests/headless/lancer.sh TestParametres        # paramètres : valeurs par défaut, modification
```

Les dialogues ne s'affichent pas sans écran : leurs contrôles sont construits,
et les clics ou saisies simulés en appelant leurs écouteurs.

## Organisation

```
calcmap/
├── build.py                    construit dist/CalcMap.oxt
├── dist/CalcMap.oxt            extension prête à installer
├── catalogue/catalogue.json    catalogue agrégé (sources de cartes), aussi inclus dans l'extension
├── src/
│   ├── basic/CalcMap.bas       code source de la macro (à modifier ici)
│   ├── CalcMap/                bibliothèque Basic (le .xba est généré par build.py)
│   ├── palettes.json           palettes de couleurs fournies
│   ├── i18n/                   textes traduits (messages_fr, messages_en…)
│   ├── description/            description de l'extension par langue
│   ├── Addons.xcu              bouton, entrée de menu, icône
│   ├── Events.xcu              reconnexion des cartes liées à l'ouverture
│   ├── CalcMap.xcs             paramètres (Catalogue, ZonesMax)
│   ├── description.xml         identifiant et version de l'extension
│   ├── META-INF/manifest.xml
│   └── icons/                  icône globe 16 et 26 px, image « pas d'aperçu »
└── tests/
    ├── generer_tests.py        (re)génère les classeurs test-<carte>.ods
    ├── test-regions.ods        13 régions        ┐
    ├── test-departements.ods   96 départements   │ Code | Valeur (test) | Nom,
    ├── test-cantons.ods        2 294 cantons     │ une ligne par zone de la carte
    ├── test-communes.ods       35 191 communes   ┘
    ├── test-cas-limites.ods    en-têtes, vides, textes, erreurs, doublons…
    ├── apercu-test-departements.png
    └── headless/               tests sans écran (lancer.sh, macros, fixtures)
```
