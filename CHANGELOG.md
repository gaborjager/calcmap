# Changelog

All notable changes to CalcMap. Each section becomes the description of the
matching GitHub release (see `.github/workflows/release.yml`).

## 0.20.2 - 2026-10-29

- Labels and values of a zone made of several parts (a country with islands
  or overseas territories, a continent) are placed on its largest part, no
  longer in the middle of all its parts.
- World maps of the ODG Atlas: continents and countries.

## 0.20.1 - 2026-10-23

- The base maps now come from the ODG Atlas project (formerly france-odg),
  which adds a map of the 27 countries of the European Union (ISO 3166-1
  alpha-2 codes, e.g. FR, DE). Map titles follow the interface language
  when the catalog translates them.

## 0.20.0 - 2026-10-23

- The project is now in English: code, comments, log messages, README and
  test tools. The user interface is still translated into 34 languages.
- Settings renamed: `Settings/Catalog` and `MaxZones` (formerly
  `Parametres/Catalogue` and `ZonesMax`). Set them again if you had changed
  them, and update Group Policy keys if you use them.
- Names stored in spreadsheets are now in English. Maps made by earlier
  versions keep working: their texts and legend are replaced on the next
  recoloring.
- The last choices offered by default (map, palette, display options) start
  again from the defaults.
- Releases are now created automatically from this changelog.

## 0.19.0 - 2026-10-23

- Narrower dialog: each column list sits below its label, and the bold,
  italic and underline buttons have their own row.

## 0.18.0 - 2026-10-23

- Separator lines in the Values and Labels frames.
- Legend frame moved to the bottom of the dialog.

## 0.17.0 - 2026-10-23

- Code, Values and Labels frames side by side, like the columns to choose.
- Icon in front of the map's example codes: a green tick if they are all in
  the chosen code column, otherwise a discreet warning whose tooltip lists the
  missing codes.

## 0.16.0 - 2026-10-23

- Separate Code frame, showing example codes of the chosen map
  (`code_examples` in the catalog).
- Column lists end with ", ..." when the column has more values.

## 0.15.1 - 2026-10-23

- More compact dialog.
- When configuring an inserted map, the map list and the insertion box are
  hidden.

## 0.15.0 - 2026-10-23

- User interface translated into 34 languages: the 24 official languages of
  the European Union, plus Arabic, Chinese, Hindi, Indonesian, Japanese,
  Korean, Norwegian, Russian, Turkish and Ukrainian.

## 0.14.1 - 2026-10-22

- Settings are changed in the Expert Configuration, or enforced by Group
  Policy on Windows; the separate configuration script was removed.

## 0.14.0 - 2026-10-22

- Clicking the CalcMap icon while an inserted map is selected opens its
  settings.
- Maps with at most 300 zones are offered by default.

## 0.13.0 - 2026-10-22

- Configurable aggregated catalog of map sources (a company can add its own).
- Maps are filtered on their number of zones (setting) instead of a fixed
  list.
- Preview of the chosen map in the dialog.

## 0.12.0 - 2026-10-22

- Outline set separately for the values and the labels.

## 0.11.0 - 2026-10-22

- Dialog organized in frames: map, code, values, labels, texts, legend.

## 0.10.0 - 2026-10-22

- Automatic update zone by zone: only the changed zones are redrawn.

## 0.9.1 - 2026-10-22

- Readability of the texts: only "with or without outline" (check box).

## 0.9.0 - 2026-10-22

- Readability of the texts: white or black outline (default) or label
  background.

## 0.8.0 - 2026-10-22

- Zone texts are drawn above the whole map, never hidden by a neighboring
  zone.
- Automatic tests without a display.

## 0.7.3 - 2026-10-19

- Classes editor: reliable number of classes when typing or using the arrows.

## 0.7.2 - 2026-10-19

- Classes editor: color bar drawn correctly (no overlapping controls).

## 0.7.1 - 2026-10-19

- Fixed: the Classes button did not open the editor.

## 0.7.0 - 2026-10-19

- Classes editor: color bar with draggable separators and color picker.

## 0.6.0 - 2026-10-19

- Custom classes, with a classes editor.

## 0.5.2 - 2026-10-16

- Column lists show the first values of each column.

## 0.5.1 - 2026-10-16

- Font, size and bold, italic, underline on one row, like the formatting
  toolbar.

## 0.5.0 - 2026-10-16

- Choice of the code, value and label columns.
- Font and style of the texts on the map.
- Legend, with or without the values.

## 0.4.4 - 2026-10-16

- Icon of the "Map settings…" context menu entry.

## 0.4.3 - 2026-10-16

- Context menu entry shown, with the extension icon.

## 0.4.1 - 2026-10-16

- Label and value centered in each zone; more robust context menu.

## 0.4.0 - 2026-10-16

- Labels (3rd column) on the map.
- Default text size 10 pt.
- Map settings by right-click on an inserted map.

## 0.3.0 - 2026-10-16

- Size of the values adjustable in the dialog (arrows or typing).

## 0.2.0 - 2026-10-15

- Option to show the values on the map.
- Cantons and municipalities maps no longer offered.

## 0.1.1 - 2026-10-15

- Fixed the automatic update of maps inserted into Calc.

## 0.1.0 - 2026-09-28

- First version: colors france-odg base maps from Calc data, with color
  palettes; the map is inserted into the sheet and updated automatically.
