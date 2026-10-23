# CalcMap - thematic maps for LibreOffice Calc

**CalcMap** is a LibreOffice extension that colors a `.odg` base map from two columns
of a Calc spreadsheet: **code | value**. It uses the base maps of the
[ODG Atlas](https://github.com/gaborjager/odg-atlas) project (France, European
Union…), where each zone carries its official code in its name (INSEE code,
ISO country code…), and any other base map built the same way.

## Usage

1. In Calc, select the data, headers included: by default code in the 1st
   column, value in the 2nd, label (optional) in the 3rd. If only two columns
   are selected, the next one is offered as the label. Whole columns can be
   selected.
2. Click the **globe** icon on the toolbar (or **Tools → CalcMap: color a map…**).
3. Choose the base map in the drop-down list (French regions and departments,
   countries of the European Union…), or another
   `.odg` file with **Other file…**.
4. Choose the color palette (a preview is shown below the list), and if needed
   customize the classes with **Classes…** (see below).
5. Check the **code**, **values** and **labels** columns: each list offers the
   selected columns, with their header and first values (followed by ", ..."
   if the column goes on) (e.g. "D - Rate: 12.5, 8.1, 9.7, ...").
6. Leave **Insert the map into the Calc sheet** checked to place the map in the
   sheet, or uncheck it to open the map in Draw.
7. Check **Show the values** and/or **Show the labels** to write the value and/or the
   label of each zone on the map, and choose their size, font and style (see
   below).
8. Check **Show the legend**, **with the values** (bounds of each class) or without.

> Only maps with at most **300 zones** are offered (`MaxZones` setting, see
> *Settings*): with the current catalog, the French regions and departments
> and the countries of the European Union. Cantons (2,306 zones) and
> municipalities (35,191) are left out.

The dialog is organized in frames: at the top **Map** (base map, its preview,
insertion in the sheet); below, side by side like the columns to choose,
**Code** (column, and examples of the map's codes to compare), **Values**
(column, colors and classes, display of the values) and **Labels** (column,
display); at the very bottom **Legend**. In front of the map's codes, a green
tick shows that they are all in the chosen code column; otherwise a small
triangle points out the missing ones (tooltip).

The map opens in Draw, colored in equal-interval classes (as many classes as
the palette has colors: 5, 7 or 10), or as a continuous gradient for the
"continuous gradient" palettes. Zones without data stay grey. The last map and
the last palette chosen are offered by default on the next run.

The user interface is translated into 34 languages (see *Languages*); the
examples below use the English one.

## Map inserted in the Calc sheet

When the box is checked (default), the colored map is placed in the sheet, to
the right of the selected data block, after an empty column, 16 cm wide. It is
anchored to its cell.

- **Up to 5,000 zones** (regions, departments): the zones are recreated in the
  sheet and grouped. Each zone keeps its color and its alternative text
  (`22 - Côtes-d'Armor - Value: 666`). To read it: double-click the map to
  enter the group, then right-click a zone → *Alt Text*.
- **Beyond that** (base map chosen with *Other file…*): the map is inserted as a
  PNG **image**. The values per zone cannot be read on the map then; uncheck
  the box to open the map in Draw if needed.

The state of the box is remembered.

### Values and labels on the map

Two boxes of the dialog write text in the center of each zone, in the inserted
map as in Draw: **Show the labels** (label on the 1st line) and **Show the values**
(value below).

- The value uses the **number format of the Calc column** (decimals,
  percentage, thousands separator…): format the values column before running
  the command. The label is used as displayed.
- The text is black on light colors, white on dark colors.
- Below each box, as on the formatting toolbar: the **font** (Liberation Sans
  by default) and the **size** (10 pt by default, from 2 to 72), chosen in the
  list or typed in, then the **B** (bold), *I* (italic) and U (underline)
  buttons. The size applies to the map inserted in Calc (16 cm wide); in Draw,
  it is adapted to keep the same look.
- A zone without a value only shows its label.
- **White or black outline** (a box under each row, checked by default): black
  text circled in white, white text circled in black, readable on all colors;
  set separately for the values and for the labels. The outline is made of
  copies of the text offset all around: a full update then takes about 3 s for
  the departments, instead of 0.6 s.

Values and labels follow the automatic update (a changed label is also taken
into account). Beyond 5,000 zones, no text is shown (too long to create and
unreadable). The texts are separate shapes, placed above all the zones: a
neighboring zone never hides them. In small zones (Paris and its inner
suburbs), they may overlap. The choices are remembered; both boxes are
unchecked by default.

### Legend

With **Show the legend** checked, a legend is placed to the right of the map: one
color box per class (5 shades for a continuous gradient), and a "No data" box
if some zones have no value.

- **with the values**: each box shows the bounds of its class (`20 – 411.2`), in
  the format of the values column, and the legend's title is the header of
  that column; otherwise, only the colors are shown.
- The text uses the font and size of the values.
- In Calc, the legend is part of the map: it moves with it and is updated with
  it.

### Settings of an inserted map

**Right-click on the map → Map settings…**, or the **globe icon** when the
map (or one of its zones) is selected, reopens the dialog for this map:
palette, columns, values, labels, legend. The map is recolored as soon as the
dialog is confirmed, and these choices (except the columns, specific to each
table) also become the starting point for the next map. The base map cannot be
changed (insert a new map instead). The menu also works on a zone, after a
double-click to enter the map.

### Automatic update

A map inserted in the sheet stays **linked to its data**: as soon as a value
changes (typed, pasted, formula result recalculated), the map is updated, with
the same settings.

- Usually, **only the changed zones** are redone (color, alternative text,
  value and label shown): about 0.3 s for one department, outline included.
- **The whole map** is redone when the change affects the whole map: new
  minimum or maximum (the linear classes move), zone that gains or loses its
  value ("No data" row of the legend), header of the values column (legend
  title), more than 20 zones changed at once, or new settings (*Map
  settings…*). For a map inserted as an image, the image is regenerated.
- The link is saved in the spreadsheet: a named range `CalcMap_1` (`_2`, … for
  other maps) on the selected columns, and a document property of the same
  name (*File → Properties → Custom Properties*). The named range follows row
  insertions; selecting whole columns (A:B) also takes rows added at the bottom
  into account.
- When the spreadsheet is opened, the extension reconnects its maps ("Open
  Document" event, declared in `Events.xcu`). Without the extension installed,
  the spreadsheet opens normally; the map simply stays frozen.
- Recoloring does not go into the undo history: *Undo* undoes the input, and
  inserting the map is a single undo step.
- To break the link: delete the map from the sheet (the link is removed on the
  next change), or delete the named range `CalcMap_N`.

> The extension uses the global "Open Document" event (*Tools → Customize →
> Events*, "Save in: LibreOffice"). If another macro was already assigned to
> it, it is replaced.

## Color palettes

Palettes provided (based on ColorBrewer and viridis):

| Type | Palettes |
|---|---|
| 5 classes | blues, greens, reds, oranges, purples, greys, yellow → red, viridis (colorblind-safe), diverging red ↔ blue and brown ↔ green |
| 7 classes | blues |
| 10 classes | blues, greens, yellow → red, viridis, diverging red ↔ blue and spectral |
| Continuous gradient | blues, yellow → red, viridis, greys, diverging red ↔ blue |

With a **continuous gradient**, there are no classes: each zone gets a color
interpolated from its exact value, between the minimum (first color) and the
maximum (last color). The palette's intermediate colors are steps of the
gradient (useful for a diverging palette: low → middle → high).

### Custom classes

By default, the values are split into equal-interval classes, one per palette
color. The **Classes…** button (next to the palette) opens the classes editor,
on the data of the chosen values column:

- **Number of classes** (arrows or typed, from 2 to 20): the classes are
  spread at equal intervals, the palette colors stretched into a gradient;
- a **color bar** scaled to the values: **dragging** a separator left or right
  moves the bound between two classes (it stays between its neighbors); the
  value of each bound is shown below; **clicking** a color opens the color
  picker;
- each box shows its **number of zones**, and under the bar one tick per value
  shows where the data are; the list on the right gives all the sorted values,
  with the label of each zone;
- **Selected bound**: the exact value of the last separator touched, to type
  on the keyboard;
- **Equal intervals** restores equal intervals, keeping the number of
  classes and the colors.

The classes always cover all the values (from the minimum to the maximum).
They are kept with the map (automatic update, legend, *Map settings…*),
but not for the next map. Choosing another palette replaces the custom
classes.

**Adding your own palettes**, without rebuilding the extension: create a
`palettes.json` file in the cache folder (see below). Its palettes are added
to the list; a palette with the same `id` as a provided one replaces it.

```json
{
  "palettes": [
    { "id": "my-palette",
      "name": { "en": "My palette", "fr": "Ma palette" },
      "colors": ["#fff5eb", "#fdae6b", "#e6550d", "#7f2704"],
      "nodata": "#ffffff" },
    { "id": "my-gradient",
      "name": "My gradient",
      "continuous": true,
      "colors": ["#ffffff", "#006d2c"] }
  ]
}
```

- `colors`: from the smallest to the largest value, at least 2 colors, as
  `#rrggbb`. Their number gives the number of classes.
- `name`: plain text, or one text per language (fallback to `en`, then to `id`).
- `nodata` (optional): color of the zones without data, light grey by default.
- `continuous` (optional): `true` for a continuous gradient; the colors are
  then the steps of the gradient and not classes.

A malformed palette is skipped and reported in the log. The extension's
palettes are in `src/palettes.json`.

## Languages

The interface follows the LibreOffice language (two-letter code). 34 languages
are provided:

- the 24 official languages of the European Union: Bulgarian (`bg`), Croatian
  (`hr`), Czech (`cs`), Danish (`da`), Dutch (`nl`), English (`en`), Estonian
  (`et`), Finnish (`fi`), French (`fr`), German (`de`), Greek (`el`), Hungarian
  (`hu`), Irish (`ga`), Italian (`it`), Latvian (`lv`), Lithuanian (`lt`),
  Maltese (`mt`), Polish (`pl`), Portuguese (`pt`), Romanian (`ro`), Slovak
  (`sk`), Slovenian (`sl`), Spanish (`es`), Swedish (`sv`);
- other widely spoken languages: Arabic (`ar`), Chinese, simplified (`zh`),
  Hindi (`hi`), Indonesian (`id`), Japanese (`ja`), Korean (`ko`), Norwegian
  Bokmål (`nb`), Russian (`ru`), Turkish (`tr`), Ukrainian (`uk`).

The extension's name, its description, the button and the menu are translated
too. English is the fallback for any untranslated text; Portuguese is European
Portuguese, Chinese uses simplified characters (also for `zh-TW`). The
translations other than French and English have not been reviewed by native
speakers: corrections are welcome.

The texts are in `src/i18n/messages_<language>.properties` (UTF-8,
`key=text`, `\n` for a line break, `%1` for an inserted value). **Adding or
fixing a language**: copy `messages_en.properties` to
`messages_<code>.properties`, translate, then rebuild the extension; `build.py`
includes the whole `i18n/` folder. For the displayed name, the description and
the menu: `src/description.xml`, `src/description/desc_<code>.txt`,
`src/Addons.xcu`. The map catalog can also provide translated fields
(`title_en`, `description_en`…), used when they exist. The log is in English.

## Base maps and cache

### Catalog

The list of maps comes from an **aggregated catalog**, read on every run: by
default [`catalogue/catalogue.json`](catalogue/catalogue.json) of this project
(path kept as is, since installed versions point to it). It lists **sources**,
each one a catalog of maps in the [ODG Atlas](https://github.com/gaborjager/odg-atlas)
format (`maps.json`); for now only one, the ODG Atlas (formerly france-odg).

```json
{
  "version": 1,
  "title": "CalcMap catalog",
  "sources": [
    { "id": "odg-atlas", "title": "ODG Atlas",
      "url": "https://raw.githubusercontent.com/gaborjager/odg-atlas/main/output/maps.json" }
  ]
}
```

- One source per catalog of maps; `id` (letters, digits, `-`, `_`) names its
  cached copy. A map whose id is already provided by a previous source is
  skipped.
- The `Catalog` setting can also point directly to a catalog of maps
  (`maps.json`).
- Accepted URLs (catalog, sources, maps, previews): `https://`, or `file:///`
  for a network share.
- Fields used for a map: `id`, `title`, `description`, `url`, `filename`,
  `size`, `sha256`, and if they exist `objects` (number of zones, for the
  filter), `preview` / `preview_sha256` (SVG, PNG or JPEG preview) and
  `code_examples` (examples of zone codes, shown in the Code frame).

### Settings

Two settings, in the LibreOffice configuration
(`/fr.github.gaborjager.CalcMap/Settings`):

| Setting | Default | Role |
|---|---|---|
| `Catalog` | `https://raw.githubusercontent.com/gaborjager/calcmap/main/catalogue/catalogue.json` | aggregated catalog (or catalog of maps) |
| `MaxZones` | `300` | maps offered: at most this number of zones |

They are changed in *Tools → Options → LibreOffice → Advanced → Open Expert
Configuration*: search for `MaxZones` or `gaborjager` (searching for "calcmap"
also finds many file paths), then double-click the value. They are taken into
account on the next run of CalcMap.

A company can provide its own aggregated catalog (its sources: the ODG Atlas, its
internal maps…). To change the sources later, publishing this catalog again at
the same address is enough.

### Deployment through Group Policy (Windows)

LibreOffice also reads its configuration from the Windows registry, under
`SOFTWARE\Policies\LibreOffice`: a Group Policy can therefore enforce the
CalcMap settings like the LibreOffice ones (*Preferences → Windows Settings →
Registry*, or an ADMX template). One key per setting, under
`HKEY_LOCAL_MACHINE` (all users of the computer):

| Key | `Value` value (REG_SZ) | `Final` value (REG_DWORD) |
|---|---|---|
| `SOFTWARE\Policies\LibreOffice\fr.github.gaborjager.CalcMap\Settings\Catalog` | `https://intranet.example.com/maps/catalog.json` | `1`: the user can no longer change it |
| `SOFTWARE\Policies\LibreOffice\fr.github.gaborjager.CalcMap\Settings\MaxZones` | `300` | `1` |

Without `Final`, the value is only a default value, which the user can change.
CalcMap must be installed (preferably for all users:
`unopkg add --shared CalcMap.oxt`); the values are read when LibreOffice
starts. To check on a computer: in the expert configuration, the enforced
value appears, and can no longer be changed if it is locked.

> Standard LibreOffice mechanism, **not tested with CalcMap** (no Windows
> computer here): to be validated on a test computer before deployment.

### Preview

The map chosen in the list is shown on the right (`preview` of the catalog,
downloaded into the cache and checked by its hash). Without a preview, or if
it is unreadable, a neutral image is shown.

### Cache

The catalog, its sources, the maps and the previews are kept in the
LibreOffice profile:

- Windows: `%APPDATA%\LibreOffice\4\user\calcmap\`
- Linux: `~/.config/libreoffice/4/user/calcmap/`

Before each use, the SHA-256 hash of the cached file is compared with the
catalog's: the map is downloaded again only if it has changed. A download
whose hash does not match is rejected. The map opens as a new document
(template): saving it never changes the cache.

**Without a network**, the cached copies of the catalog and of the sources are
used, and the list only offers the maps already downloaded (and any other
`.odg` placed in this folder). On the very first launch without a network, the
catalog shipped with the extension is the fallback. Network access goes
through LibreOffice: a proxy is set in **Tools → Options → Internet → Proxy**.

## Handling imperfect data

The macro never shows a Basic error. Unusable data leaves the zone in the
default color (grey) and is noted in the log.

| Case | Handling |
|---|---|
| Empty row, empty code | skipped |
| Header (even repeated in the middle) | skipped (non-numeric value) |
| Empty value, text (`n.d.`), cell in error (`#DIV/0!`) | default color |
| Number typed as text (`1 234,5`) | accepted |
| Code without its leading zero (`1` for `01`) | recovered |
| Code in lowercase or with spaces (` 2a `) | accepted |
| Duplicate code | last value kept |
| Code not in the map | skipped |

## Alternative text

The value of each zone is added to its **alternative text** (right-click a
zone → *Alt Text*, *Text Alternative* field), after the text provided by the
map:

- `22 - Côtes-d'Armor` becomes `22 - Côtes-d'Armor - Value: 666`
- without data: `22 - Côtes-d'Armor - Value: (none)`

The words follow the interface language. Recoloring an already colored map
replaces the value instead of adding it a second time. Can be disabled with
`WRITE_ALT_TEXT = False` in the code.

## Log (debugging)

Each run overwrites the `calcmap.log` file of the LibreOffice profile:

- Windows: `%APPDATA%\LibreOffice\4\user\calcmap.log`
- Linux: `~/.config/libreoffice/4/user/calcmap.log`

It lists the language, the palettes, the range read, the catalog and the cache
(hashes, downloads), each discarded row with its reason, the codes unknown to
the map and the summary. With `DETAILED_LOG = True`, it also traces each zone.

## Installation

Download `CalcMap.oxt` from the latest [release](https://github.com/gaborjager/calcmap/releases/latest)
(what changed: [CHANGELOG.md](CHANGELOG.md)), or the latest build
[`dist/CalcMap.oxt`](https://github.com/gaborjager/calcmap/raw/main/dist/CalcMap.oxt), then:

- **One user**: double-click `CalcMap.oxt`, or **Tools → Extension Manager →
  Add**, then restart LibreOffice.
- **All users of a computer** (administrator, LibreOffice closed):
  ```
  "C:\Program Files\LibreOffice\program\unopkg.exe" add --shared CalcMap.oxt
  ```

The button is added to the Calc *Standard* toolbar, after *Insert Chart*. With
the tabbed interface, use **Tools → CalcMap: color a map…**.

### Upgrading from version 0.19 or earlier

Since version 0.20, the code, the settings and the names stored in documents
are in English. Existing spreadsheets keep working: the zone texts and legends
created by earlier versions ("CalcMap texte", "CalcMap légende") are
recognized and replaced on the next recoloring. Two things are not carried
over:

- the **settings** were renamed (`Parametres/Catalogue` and `ZonesMax` are now
  `Settings/Catalog` and `MaxZones`): set them again if you had changed them
  (or update your Group Policy keys);
- the **last choices** (map, palette, display options offered by default) start
  again from the defaults.

### Upgrading from ColorierCarte

CalcMap is the new name of the *ColorierCarte* extension. It is a new extension
for LibreOffice: **uninstall ColorierCarte** in the Extension Manager before
installing CalcMap, otherwise both buttons coexist. Nothing is lost:

- on first launch, the `coloriercarte/` cache (downloaded maps, personal
  `palettes.json` palettes) is copied into `calcmap/`;
- spreadsheets whose map was linked by ColorierCarte (`ColorierCarte_N` ranges)
  keep updating; new maps use `CalcMap_N`.

## Building the extension

```bash
python3 build.py
```

Produces `dist/CalcMap.oxt`. The version is read from the `VERSION` constant of
the Basic code and copied into `description.xml`.

### Releases

Releases are created by GitHub Actions (`.github/workflows/release.yml`): on
every push to `main`, if the tag `v<VERSION>` does not exist yet, the
extension is built and the release "CalcMap <VERSION>" is created, described
by the matching section of [CHANGELOG.md](CHANGELOG.md), with `CalcMap.oxt`
attached. To publish a version: raise `VERSION`, add its section at the top
of `CHANGELOG.md`, commit and push. The workflow fails if the section is
missing (`python3 tools/release_notes.py notes <version>` shows it).

The repository needs *Settings → Actions → General → Workflow permissions →
Read and write permissions* for the workflow to create releases.

## Test files

One spreadsheet per base map of the catalog (the cantons and municipalities
ones are used if the `MaxZones` setting is raised): `tests/test-<map>.ods`,
with the columns *Code | Value (test) | Name* and one row per zone. Select
columns A:B then run the command with the matching map: all the zones must be
colored, without unknown codes in the log.

To regenerate them from the catalog's maps (random but reproducible values):

```bash
python3 tests/generate_tests.py                        # missing files
python3 tests/generate_tests.py departements --force   # rewrite one map
```

### Automatic tests (without a display)

`tests/headless/run.sh <macro>` runs a test macro of
`tests/headless/tests_calcmap.bas` in LibreOffice without a display, on a copy
of the extension and a throwaway profile, and prints the `TEST` lines of the
log:

```bash
tests/headless/run.sh TestTexts            # zone texts, legend
tests/headless/run.sh TestFullEditor       # classes editor (mouse, arrows, typing)
tests/headless/run.sh TestPartial          # zone-by-zone update vs full recalculation
tests/headless/run.sh TestDialog           # main dialog: settings read back unchanged
tests/headless/run.sh TestCodesIcon        # code check icon, dialog geometry
tests/headless/run.sh TestCatalog          # aggregated catalog, filter, previews (local files)
tests/headless/run.sh TestIconOnMap        # icon on a selected map: configuration
tests/headless/run.sh TestSettings         # settings: default values, change
tests/headless/run.sh TestTranslations     # each language loaded; dialog in de, ar, ja
tests/headless/run.sh TestOldNames         # maps made by versions up to 0.19
tests/headless/run.sh TestEuMap            # European Union map colored from ISO codes
```

The tests use a local copy of the ODG Atlas, next to this repository
(`../odg-atlas`, or the path in the `ODG_ATLAS` environment variable).

Dialogs are not shown without a display: their controls are built, and clicks
or typing are simulated by calling their listeners.

## Layout

```
calcmap/
├── build.py                    builds dist/CalcMap.oxt
├── CHANGELOG.md                changes per version (release descriptions)
├── .github/workflows/          release.yml: GitHub release for each new version
├── tools/release_notes.py      version and release notes for the workflow
├── dist/CalcMap.oxt            extension ready to install
├── catalogue/catalogue.json    aggregated catalog (map sources), also included in the extension
├── src/
│   ├── basic/CalcMap.bas       source code of the macro (edit here)
│   ├── CalcMap/                Basic library (the .xba is generated by build.py)
│   ├── palettes.json           color palettes provided
│   ├── i18n/                   translated texts (34 languages: messages_en, messages_fr…)
│   ├── description/            description of the extension per language
│   ├── Addons.xcu              button, menu entry, icon
│   ├── Events.xcu              reconnection of the linked maps on opening
│   ├── CalcMap.xcs             settings (Catalog, MaxZones)
│   ├── description.xml         id and version of the extension
│   ├── META-INF/manifest.xml
│   └── icons/                  globe icon 16 and 26 px, "no preview" image, code check icons
└── tests/
    ├── generate_tests.py       (re)generates the test-<map>.ods spreadsheets
    ├── test-regions.ods        13 regions        ┐
    ├── test-departements.ods   96 departments    │ Code | Value (test) | Name,
    ├── test-cantons.ods        2,294 cantons     │ one row per zone of the map
    ├── test-communes.ods       35,191 communes   ┘
    ├── test-edge-cases.ods     headers, blanks, texts, errors, duplicates…
    ├── preview-test-departements.png
    └── headless/               tests without a display (run.sh, driver.py, macros, fixtures)
```
