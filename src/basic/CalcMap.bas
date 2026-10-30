REM  CalcMap - thematic maps for LibreOffice Calc
REM  ============================================
REM  Colors the zones of a .odg base map (one shape per zone, named with
REM  its code, e.g. INSEE code) from 2 Calc columns: code | value.
REM
REM  Usage:
REM   1. In Calc, select the data: by default code in the 1st column,
REM      value in the 2nd, optional label in the 3rd; the dialog lets you
REM      choose other columns of the selection.
REM   2. Click the globe icon (or Tools > CalcMap: color a map…).
REM   3. Choose the base map in the list (ODG Atlas catalog), or another
REM      .odg file with "Other file…", and the color palette.
REM
REM  Base maps: the aggregated catalog (Catalog setting, by default the
REM  project's catalogue/catalogue.json) lists sources, each one a catalog
REM  of maps in the ODG Atlas format (maps.json); it is read on every
REM  run. Only maps with at most MaxZones zones (setting, 300 by default)
REM  are offered. Maps are downloaded into the LibreOffice profile
REM  ($(user)/calcmap/) and downloaded again only if their SHA-256
REM  hash differs from the catalog's. Without a network, only maps
REM  already downloaded are offered.
REM
REM  Insertion into Calc (dialog option): the colored map is placed in the
REM  sheet, to the right of the data. Up to CALC_SHAPES_LIMIT zones, the
REM  shapes are recreated in the sheet (each zone keeps its name and its
REM  alternative text); beyond that (municipalities), the map is inserted as an image.
REM
REM  Automatic update: a map inserted into Calc stays linked to its data
REM  (named range CalcMap_N + document property of the same name). Any
REM  change to the values recolors the map. The link is saved with the
REM  spreadsheet and reconnected when it is opened ("Open Document" event
REM  declared by the extension in Events.xcu).
REM
REM  Values and labels on the map (dialog options): the label and the value
REM  of each zone are written in a separate text shape (named TEXT_PREFIX),
REM  centered on the zone and placed after all the zones so that it is
REM  never hidden by a neighboring zone; two paragraphs, the value in the number format of the Calc
REM  column, white on dark colors and black on light ones, at the
REM  sizes chosen in the dialog (points, on the map inserted into Calc;
REM  same look in Draw). Beyond CALC_SHAPES_LIMIT
REM  zones, neither values nor labels are shown.
REM
REM  Legend (option): one color box per class, with or without its bounds,
REM  to the right of the map and in its group (shapes named LEGEND_PREFIX),
REM  recreated on every recoloring.
REM
REM  Configuration: right-click on an inserted map -> "Map
REM  settings…" (intercepted context menu, InterceptMenu) reopens the dialog
REM  to change the palette, values and labels of this map.
REM
REM  Palettes: palettes.json of the extension, completed (or replaced, for
REM  the same id) by $(user)/calcmap/palettes.json. The number of colors
REM  of a palette sets the number of classes; a "continuous" palette
REM  gives a continuous gradient (color interpolated from the exact value).
REM
REM  Languages: i18n/messages_<language>.properties of the extension,
REM  following the interface language; English is the key-by-key fallback.
REM  The log is in English.
REM
REM  Robustness: the macro must never show a Basic error.
REM   - empty rows, headers, empty codes: skipped
REM   - missing or empty value, non-numeric text, cell in error:
REM     the zone keeps the default color (grey)
REM   - numbers typed as text ("1 234,5"): accepted
REM   - leading zeros lost by Calc (01 -> 1): recovered
REM
REM  Log: each run overwrites the calcmap.log file in the
REM  LibreOffice profile (Windows: %APPDATA%\LibreOffice\4\user\,
REM  Linux: ~/.config/libreoffice/4/user/).
REM  Set DETAILED_LOG to True to also trace each zone.

Option Explicit

Const VERSION = "0.21.0"
Const ID_EXTENSION = "fr.github.gaborjager.calcmap"
Const DETAILED_LOG = False              ' True: one line per zone in the log
Const WRITE_ALT_TEXT = True             ' True: value added to the alternative text (Description)
Const LOG_NAME = "calcmap.log"
' Settings (CalcMap.xcs): catalog and maximum number of zones; fallback
' values if the configuration is unreadable
Const SETTINGS_NODE = "/fr.github.gaborjager.CalcMap/Settings"
Const DEFAULT_CATALOG = "https://raw.githubusercontent.com/gaborjager/calcmap/main/catalogue/catalogue.json"
Const DEFAULT_MAX_ZONES = 300
Const MAPS_FOLDER = "calcmap"           ' cache of the base maps, in the profile
Const OLD_FOLDER = "coloriercarte"      ' cache of the former ColorierCarte extension (taken over)
Const CATALOG_NAME = "catalog.json"     ' local copy of the aggregated catalog (offline mode)
Const PREVIEWS_FOLDER = "previews"      ' map previews, in the cache
Const DEFAULT_MAP = "departements"
Const PALETTES_NAME = "palettes.json"   ' palettes of the extension and of the user
Const PREFERENCES_NAME = "preferences.properties" ' last chosen map and palette
Const DEFAULT_PALETTE = "bleus"
Const PREVIEW_SHADES = 12               ' palette preview boxes in the dialog
Const PREVIEW_X = 164                   ' position and width of the preview (dialog units)
Const PREVIEW_W = 132
Const CALC_SHAPES_LIMIT = 5000          ' beyond this, map inserted into Calc as an image
Const CALC_MAP_WIDTH = 16000            ' width of the inserted map (1/100 mm)
Const TEXT_FONT = "Liberation Sans"     ' values and labels on the map (shipped with LibreOffice)
Const TEXT_SIZE = 10                    ' default size of values and labels (points)
Const TEXT_SIZE_MIN = 2
Const TEXT_SIZE_MAX = 72
Const OFFERED_SIZES = "6|7|8|9|10|11|12|14|16|18|20|24|28|32|36|48|72" ' list of sizes in the dialog
Const IMAGE_WIDTH_PX = 2400             ' width of the PNG image (detailed maps)
Const IMAGE_NAME = "inserted-map.png"   ' temporary image, in the cache
Const CONFIGURE_URL = "vnd.sun.star.script:CalcMap.CalcMap.ConfigureMap?language=Basic&location=application"
Const LEGEND_PREFIX = "CalcMap legend"  ' name of the legend shapes
Const TEXT_PREFIX = "CalcMap text"      ' name of the zone texts (followed by the code)
' Names and values written by versions up to 0.19 (French), still recognized
' in existing spreadsheets; "CalcMap texte" already starts with TEXT_PREFIX
Const OLD_LEGEND_PREFIX = "CalcMap légende"
Const OLD_TEXT_PREFIX = "CalcMap texte"
Const OLD_OUTLINE_NONE = "aucune"
Const PREVIEW_VALUES = 3                ' values shown per column in the dialog
Const PREVIEW_ROWS = 20                 ' rows scanned to find them
Const MAX_CLASSES = 20                  ' maximum number of custom classes
Const MAX_COLUMNS = 26                  ' maximum number of columns offered in the dialog
Const LINK_PREFIX = "CalcMap_"          ' named ranges and properties of the linked maps
Const OLD_LINK_PREFIX = "ColorierCarte_" ' links created by the former extension (recognized)

' Fields of a catalog map (array returned by NewMap)
Const C_ID = 0
Const C_TITLE = 1
Const C_DESCRIPTION = 2
Const C_URL = 3
Const C_FILE = 4
Const C_SIZE = 5
Const C_SHA256 = 6
Const C_LOCALE = 7
Const C_OBJECTS = 8                     ' number of zones (0: unknown)
Const C_PREVIEW = 9                  ' preview URL ("": none)
Const C_PREVIEW_SHA = 10
Const C_SOURCE = 11                     ' source in the aggregated catalog
Const C_CODES = 12                  ' examples of zone codes, e.g. "01, 02, 03, 04" ("": unknown)

' Options of a map (array returned by DefaultOptions). Same order in
' the link (from the 6th field on); names in the preferences: OPTION_NAMES.
Const O_VALUES = 0
Const O_SIZE_VALUES = 1
Const O_LABELS = 2
Const O_SIZE_LABELS = 3
Const O_FONT_VALUES = 4
Const O_STYLE_VALUES = 5                ' letters B (bold), I (italic), U (underline)
Const O_FONT_LABELS = 6
Const O_STYLE_LABELS = 7
Const O_COL_CODE = 8                    ' columns of the range (0: 1st); specific to each map
Const O_COL_VALUE = 9
Const O_COL_LABEL = 10
Const O_LEGEND = 11
Const O_LEGEND_VALUES = 12
Const O_CLASSES = 13                ' custom classes (ClassesText); "": linear distribution
Const O_OUTLINE_VALUES = 14             ' outline of the values: OUTLINE_ON or OUTLINE_NONE
Const O_OUTLINE_LABELS = 15             ' outline of the labels: same
Const OPTION_COUNT = 16
Const OPTION_NAMES = "values|value_size|labels|label_size|value_font|value_style|label_font|label_style|code_col|value_col|label_col|legend|legend_values|classes|value_outline|label_outline"
Const OPTION_TYPES = "bnbnsssscccbbkss" ' b: yes/no, n: size, s: text, c: column, k: classes
Const OUTLINE_NONE = "none"
Const OUTLINE_ON = "outline"            ' outline in the opposite color (offset copies of the text)
' Options c and k, specific to each table, are kept in the map's link
' but not in the preferences.

' Context of a coloring (array returned by TextContext)
Const T_FORMATTER = 0
Const T_FORMAT = 1
Const T_OPTIONS = 2
Const T_DOCUMENT = 3                    ' document where the legend is created (Nothing: no legend)
Const T_TITLE = 4                       ' legend title: header of the values column

' Fields of a palette (array returned by NewPalette)
Const P_ID = 0
Const P_NAME = 1
Const P_COLORS = 2
Const P_EMPTY = 3
Const P_CONTINUOUS = 4                  ' True: continuous gradient, otherwise classes

Private oLog As Object                  ' text stream of the log (Null if unavailable)
Private sLogPath As String
Private oDialog As Object               ' map choice dialog (for its listeners)
Private bKeepDialog As Boolean          ' tests: dialog kept after execute(), to read it back
Private aDialogMaps As Variant
Private aDialogPalettes As Variant
Private aDialogData As Variant          ' data of the selection (classes editor)
Private sDialogClasses As String  ' current custom classes ("": linear)
Private oEditor As Object               ' classes editor and its state
Private aEdBounds As Variant
Private aEdColors As Variant
Private nEdClasses As Integer
Private aEdValues As Variant            ' sorted values of the values column
Private aEdPalette As Variant
Private bEdUpdating As Boolean          ' fields being updated (events ignored)
Private dBarMin As Double               ' values at the ends of the editor's bar
Private dBarMax As Double
Private iEdBound As Integer             ' selected bound (-1: none)
Private bEdDragging As Boolean          ' separator being dragged
Private iEdColor As Integer             ' clicked class (its color is chosen on release)
Private nEdShown As Integer             ' classes shown on the bar (-1: to recompute)

' Color bar of the classes editor (dialog units)
Const BAR_X = 10
Const BAR_Y = 44
Const BAR_W = 300
Const BAR_H = 22
Const PARTIAL_MAX_ZONES = 20            ' beyond this, full update rather than zone by zone
Const MAX_TICKS = 300                   ' maximum number of value ticks under the bar
Private oTexts As Object                ' translated texts (key -> text), loaded on first use
Private sLanguage As String
' Linked maps: Global and not Private, because these variables must outlive
' the end of the macro (the modify listener is called later)
Global aLinks() As Variant              ' watched linked maps: Array(document, range, listener, name, memory)
Global nLinks As Long
Global bRecalcRunning As Boolean
Global nLastUpdate As Long              ' zones redone at the last update (-1: whole map)
Global aMenus() As Variant              ' intercepted context menus: Array(controller, interceptor)
Global nMenus As Long
Private sJson As String                 ' state of the JSON reader
Private nJson As Long
Private bJsonOk As Boolean

' ---------------------------------------------------------------------------
' Entry point (button, menu)
' ---------------------------------------------------------------------------
Sub CalcMap()
	Dim oDocCalc As Object, oRange As Object, aData As Variant
	Dim sUrl As String, oDoc As Object, nOk As Long, bFromCache As Boolean, aPalette As Variant
	Dim bInsert As Boolean, oSelection As Object, oShp As Object, bImage As Boolean
	Dim aOptions As Variant, vText As Variant, aColumns As Variant, sLink As String
	Dim oLegendDoc As Object
	Dim aProps(1) As New com.sun.star.beans.PropertyValue
	On Error GoTo Failed

	OpenLog()
	WriteLog("CalcMap " & VERSION & " - start")
	LoadTexts()

	oDocCalc = ThisComponent
	If IsNull(oDocCalc) Then
		Warn(Tr("msg.open_calc"))
		GoTo Finish
	End If
	If Not oDocCalc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
		Warn(Tr("msg.use_from_calc"))
		GoTo Finish
	End If

	' Linked map selected (or one of its zones): configure this map
	' rather than create a new one
	If SelectionLink(oDocCalc, oDocCalc.getCurrentController().getSelection()) <> "" Then
		WriteLog("Linked map selected: configuring this map")
		CloseLog()
		ConfigureMap()
		Exit Sub
	End If

	oSelection = SelectedRange(oDocCalc.getCurrentSelection())
	oRange = DataRange(oSelection)
	If IsNull(oRange) Then
		Warn(Tr("msg.select_two_columns"))
		GoTo Finish
	End If
	WriteLog("Range read: " & RangeName(oRange))
	aData = oRange.getDataArray()
	aColumns = ColumnNames(oRange)

	sUrl = ChooseBaseMap(oDocCalc, bFromCache, aPalette, bInsert, aOptions, aColumns, aData)
	If sUrl = "" Then
		WriteLog("Base map choice canceled or impossible.")
		GoTo Finish
	End If
	WriteLog("Base map: " & ConvertFromURL(sUrl))

	' A cached map opens as a template: saving it does not change the cache.
	' For an insertion into Calc, the drawing stays hidden and is then closed.
	aProps(0).Name = "AsTemplate"
	aProps(0).Value = bFromCache
	aProps(1).Name = "Hidden"
	aProps(1).Value = bInsert
	oDoc = StarDesktop.loadComponentFromURL(sUrl, "_blank", 0, aProps())
	If IsNull(oDoc) Then
		Warn(Tr("msg.cannot_open_map"))
		GoTo Finish
	End If
	If Not oDoc.supportsService("com.sun.star.drawing.DrawingDocument") Then
		Warn(Tr("msg.not_drawing"))
		GoTo Finish
	End If

	If (aOptions(O_VALUES) Or aOptions(O_LABELS)) And ZoneCount(oDoc) > CALC_SHAPES_LIMIT Then
		' Tens of thousands of texts: several minutes, and unreadable
		WriteLog("Too many zones (" & ZoneCount(oDoc) & "): neither values nor labels")
		Inform(Tr("msg.too_many_values", ZoneCount(oDoc)))
		aOptions(O_VALUES) = False
		aOptions(O_LABELS) = False
	End If
	' Texts and legend: created in Draw; for a map inserted into Calc,
	' by the recoloring that follows the insertion (only the zones are copied)
	If Not bInsert Then oLegendDoc = oDoc
	vText = TextContext(oDocCalc, oRange, aData, aOptions, oLegendDoc)
	nOk = ColorDocument(oDoc, aData, aPalette, vText)
	If nOk = 0 Then Warn(Tr("msg.no_match"))
	If bInsert Then
		If nOk > 0 Then
			' Insertion and link = a single action in the undo history
			oDocCalc.getUndoManager().enterUndoContext(Tr("undo.insert_map"))
			oShp = InsertIntoCalc(oDoc, oDocCalc, oRange, FileNameOf(sUrl), bImage)
			If Not IsNull(oShp) Then
				sLink = CreateLink(oDocCalc, oSelection, oShp, aPalette(P_ID), sUrl, bImage, aOptions)
				' Texts and legend: created by the recoloring in the sheet
				If (aOptions(O_LEGEND) Or aOptions(O_VALUES) Or aOptions(O_LABELS)) And sLink <> "" Then
					RecalcLink(oDocCalc, sLink)
				End If
			End If
			oDocCalc.getUndoManager().leaveUndoContext()
		End If
		oDoc.close(True)
	End If
	GoTo Finish

Failed:
	WriteLog("Unexpected ERROR " & Err & " (line " & Erl & "): " & Error$)
	On Error Resume Next
	Warn(Tr("msg.partial"))
	Resume Finish
Finish:
	On Error Resume Next
	WriteLog("Done.")
	CloseLog()
End Sub

' ---------------------------------------------------------------------------
' Core: colors the shapes of an open Draw document with a palette:
' as many equal-interval classes as colors, or a continuous gradient
' between the palette colors. Returns the number of colored zones.
' vText: context of the coloring (TextContext): columns, values,
' labels, legend.
' Never raises an error.
' ---------------------------------------------------------------------------
Function ZoneCount(oDoc As Object) As Long
	Dim j As Long
	On Error Resume Next
	For j = 0 To oDoc.getDrawPages().getCount() - 1
		ZoneCount = ZoneCount + oDoc.getDrawPages().getByIndex(j).getCount()
	Next j
End Function

Function ColorDocument(oDoc As Object, aData As Variant, aPalette As Variant, vText As Variant) As Long
	Dim aPages() As Object, j As Long
	On Error GoTo Failed
	ReDim aPages(oDoc.getDrawPages().getCount() - 1) As Object
	For j = 0 To UBound(aPages)
		aPages(j) = oDoc.getDrawPages().getByIndex(j)
	Next j
	ColorDocument = ColorContainers(aPages, aData, aPalette, vText)
	Exit Function
Failed:
	WriteLog("ERROR while coloring: " & Error$)
	ColorDocument = 0
End Function

' Colors the shapes of each container (draw page, group…)
Function ColorContainers(aContainers As Variant, aData As Variant, aPalette As Variant, _
		vText As Variant) As Long
	Dim oMap As Object, dMin As Double, dMax As Double, nValid As Long
	Dim oLabels As Object, aOptions As Variant, bText As Boolean, dScale As Double, nZones As Long
	Dim dValueSize As Double, dLabelSize As Double, sValue As String, sLabel As String
	Dim aColors As Variant, nClasses As Integer, oPage As Object, oShape As Object
	Dim bContinuous As Boolean, dPosition As Double, nColor As Long
	Dim sName As String, dVal As Double, iClass As Integer
	Dim i As Long, j As Long, nOk As Long, nWithout As Long
	Dim oUsed As Object, sKey As String, oEnum As Object, nUnknown As Long
	Dim bCustom As Boolean, aBounds As Variant, aCustomColors As Variant, nOutside As Long
	Dim aZonesT As Variant, aLabT As Variant, aValT As Variant, aClrT As Variant, nT As Long
	On Error GoTo Failed
	ColorContainers = 0
	aZonesT = Array() : aLabT = Array() : aValT = Array() : aClrT = Array()
	oUsed = com.sun.star.container.EnumerableMap.create("string", "boolean")
	aOptions = vText(T_OPTIONS)
	bCustom = ReadClasses(aOptions(O_CLASSES), aBounds, aCustomColors)

	oMap = com.sun.star.container.EnumerableMap.create("string", "double")
	oLabels = com.sun.star.container.EnumerableMap.create("string", "string")
	nValid = ReadData(aData, oMap, dMin, dMax, oLabels, aOptions(O_COL_CODE), _
		aOptions(O_COL_VALUE), aOptions(O_COL_LABEL))
	If nValid = 0 Then
		WriteLog("No usable data: the map keeps the default color.")
	Else
		WriteLog("Min = " & dMin & " ; max = " & dMax)
	End If

	aColors = aPalette(P_COLORS)
	nClasses = UBound(aColors) + 1
	bContinuous = aPalette(P_CONTINUOUS)
	If bCustom Then
		WriteLog("Custom classes: " & aOptions(O_CLASSES))
	ElseIf bContinuous Then
		WriteLog("Palette: " & aPalette(P_ID) & " (continuous gradient, " & nClasses & " colors)")
	Else
		WriteLog("Palette: " & aPalette(P_ID) & " (" & nClasses & " classes)")
	End If

	' Text sizes (0: not shown): the ones chosen for the map inserted
	' into Calc, proportional to the map width to keep the same look in
	' Draw. The zone text is always handled (a disabled text is
	' cleared), except beyond CALC_SHAPES_LIMIT zones.
	For j = 0 To UBound(aContainers)
		nZones = nZones + aContainers(j).getCount()
	Next j
	dScale = ZonesWidth(aContainers(0)) / CALC_MAP_WIDTH
	If nZones > CALC_SHAPES_LIMIT Then
		WriteLog("Too many zones (" & nZones & "): neither values nor labels")
	Else
		bText = True
		If aOptions(O_VALUES) Then dValueSize = aOptions(O_SIZE_VALUES) * dScale
		If aOptions(O_LABELS) Then dLabelSize = aOptions(O_SIZE_LABELS) * dScale
	End If
	WriteLog("Options: " & OptionsText(aOptions))

	For j = 0 To UBound(aContainers)
		oPage = aContainers(j)
		For i = 0 To oPage.getCount() - 1
			oShape = oPage.getByIndex(i)
			sName = ""
			On Error Resume Next
			sName = oShape.getName()
			On Error GoTo Failed
			If sName <> "" And Not IsAddOn(sName) Then
				sLabel = ""
				If dLabelSize > 0 Then sLabel = FindLabel(oLabels, sName)
				If FindValue(oMap, sName, dVal, sKey) Then
					oUsed.put(sKey, True)
					nColor = ValueColor(dVal, dMin, dMax, aPalette, bCustom, aBounds, aCustomColors)
					If nColor < 0 Then
						nColor = aPalette(P_EMPTY)
						nOutside = nOutside + 1
						WriteLog("  " & sName & " = " & dVal & ": outside the classes -> default color")
					End If
					sValue = ""
					If dValueSize > 0 Then sValue = ValueText(vText, dVal)
					ApplyZone(oShape, nColor, Tr("alt.value", dVal), bText)
					KeepText(oShape, sLabel, sValue, nColor, aZonesT, aLabT, aValT, aClrT, nT)
					nOk = nOk + 1
					If DETAILED_LOG Then WriteLog("  " & sName & " = " & dVal & " -> color " & Hex(nColor))
				Else
					ApplyZone(oShape, aPalette(P_EMPTY), Tr("alt.none"), bText)
					KeepText(oShape, sLabel, "", aPalette(P_EMPTY), aZonesT, aLabT, aValT, aClrT, nT)
					nWithout = nWithout + 1
					If DETAILED_LOG Then WriteLog("  " & sName & ": no data -> default color")
				End If
			End If
		Next i
	Next j

	' Codes present in the data but missing from the map
	oEnum = oMap.createKeyEnumeration(False)
	Do While oEnum.hasMoreElements()
		sKey = oEnum.nextElement()
		If Not oUsed.containsKey(sKey) Then
			WriteLog("Code '" & sKey & "' not in the base map -> skipped")
			nUnknown = nUnknown + 1
		End If
	Loop
	WriteLog("Zones colored: " & nOk & " ; no data (default color): " & nWithout _
		& " ; codes unknown to the map: " & nUnknown & IIf(bCustom, " ; outside the classes: " & nOutside, ""))

	' Zone texts and legend: removed then recreated, above the zones
	If Not IsNull(vText(T_DOCUMENT)) Then
		RemoveAddOns(aContainers(0))
		If bText Then
			For i = 0 To nT - 1
				ZoneText(vText(T_DOCUMENT), aContainers(0), aZonesT(i), aLabT(i), dLabelSize, aValT(i), _
					dValueSize, aOptions, aClrT(i))
			Next i
		End If
		If aOptions(O_LEGEND) Then CreateLegend(vText, aContainers(0), aPalette, dMin, dMax, _
			nValid > 0, nWithout + nOutside > 0, dScale)
	End If
	ColorContainers = nOk
	Exit Function
Failed:
	WriteLog("ERROR while coloring " & Err & " (line " & Erl & "): " & Error$)
	ColorContainers = nOk
End Function

' Color of a value: custom class, continuous gradient or equal-interval
' class between dMin and dMax. -1: outside the custom classes.
Function ValueColor(ByVal dVal As Double, ByVal dMin As Double, ByVal dMax As Double, aPalette As Variant, _
		bCustom As Boolean, aBounds As Variant, aCustomColors As Variant) As Long
	Dim dPosition As Double, iClass As Integer, nClasses As Integer, aColors As Variant
	If bCustom Then
		iClass = ClassOf(aBounds, dVal)
		If iClass < 0 Then ValueColor = -1 Else ValueColor = aCustomColors(iClass)
		Exit Function
	End If
	' Position of the value between the min (0) and the max (1)
	If dMax = dMin Then dPosition = 1 Else dPosition = (dVal - dMin) / (dMax - dMin)
	aColors = aPalette(P_COLORS)
	If aPalette(P_CONTINUOUS) Then
		ValueColor = GradientColor(aColors, dPosition)
	Else
		nClasses = UBound(aColors) + 1
		iClass = Int(dPosition * nClasses)
		If iClass > nClasses - 1 Then iClass = nClasses - 1
		If iClass < 0 Then iClass = 0
		ValueColor = aColors(iClass)
	End If
End Function

' Gradient color at position dPosition (0 to 1), interpolated
' linearly between the successive colors of the palette
Function GradientColor(aColors As Variant, ByVal dPosition As Double) As Long
	Dim nSegments As Integer, i As Integer, f As Double, c1 As Long, c2 As Long
	If dPosition < 0 Then dPosition = 0
	If dPosition > 1 Then dPosition = 1
	nSegments = UBound(aColors)
	i = Int(dPosition * nSegments)
	If i > nSegments - 1 Then i = nSegments - 1
	f = dPosition * nSegments - i
	c1 = aColors(i)
	c2 = aColors(i + 1)
	GradientColor = RGB(Int(Red(c1) + (Red(c2) - Red(c1)) * f + 0.5), _
		Int(Green(c1) + (Green(c2) - Green(c1)) * f + 0.5), _
		Int(Blue(c1) + (Blue(c2) - Blue(c1)) * f + 0.5))
End Function

' Applies color and alternative text, and if bText the zone text
' (label and value; empty: text cleared); a stubborn shape is
' skipped
Sub ApplyZone(oShape As Object, nColor As Long, sText As String, bText As Boolean)
	On Error GoTo Failed
	oShape.FillStyle = com.sun.star.drawing.FillStyle.SOLID
	oShape.FillColor = nColor
	If WRITE_ALT_TEXT Then oShape.Description = AltText(oShape.Description, sText)
	' Text written in the zone by previous versions: removed (texts
	' are now separate shapes, above the map)
	If bText Then
		If oShape.getString() <> "" Then oShape.setString("")
	End If
	Exit Sub
Failed:
	WriteLog("  shape skipped (" & Error$ & ")")
End Sub

' Keeps the text of a zone (label and/or value), created after the coloring
Sub KeepText(oZone As Object, sLabel As String, sValue As String, nColor As Long, _
		aZones As Variant, aLab As Variant, aVal As Variant, aClr As Variant, n As Long)
	If sLabel = "" And sValue = "" Then Exit Sub
	ReDim Preserve aZones(n)
	ReDim Preserve aLab(n)
	ReDim Preserve aVal(n)
	ReDim Preserve aClr(n)
	aZones(n) = oZone
	aLab(n) = sLabel
	aVal(n) = sValue
	aClr(n) = nColor
	n = n + 1
End Sub

' Text shape of a zone: label then value, centered on the zone, added
' to the container after the zones (so above them); color depends on the background
Sub ZoneText(oDoc As Object, oContainer As Object, oZone As Object, sLabel As String, dLabelSize As Double, _
		sValue As String, dValueSize As Double, aOptions As Variant, nBackground As Long)
	Dim aTexts(1) As String, aSizes(1) As Double, aFonts(1) As String, aStyles(1) As String
	Dim aOutlines(1) As Boolean, bOutline As Boolean, aCopies(1) As String
	Dim n As Integer, k As Integer, dCx As Double, dCy As Double
	Dim nColor As Long, dGap As Double, aDx As Variant, aDy As Variant, oText As Object
	On Error GoTo Failed
	If sLabel <> "" Then
		aTexts(n) = sLabel : aSizes(n) = dLabelSize
		aFonts(n) = aOptions(O_FONT_LABELS) : aStyles(n) = aOptions(O_STYLE_LABELS)
		aOutlines(n) = HasOutline(aOptions(O_OUTLINE_LABELS))
		bOutline = bOutline Or aOutlines(n)
		n = n + 1
	End If
	If sValue <> "" Then
		aTexts(n) = sValue : aSizes(n) = dValueSize
		aFonts(n) = aOptions(O_FONT_VALUES) : aStyles(n) = aOptions(O_STYLE_VALUES)
		aOutlines(n) = HasOutline(aOptions(O_OUTLINE_VALUES))
		bOutline = bOutline Or aOutlines(n)
		n = n + 1
	End If
	If n = 0 Then Exit Sub
	ReDim Preserve aTexts(n - 1) As String
	ZoneCenter(oZone, dCx, dCy)
	nColor = ContrastTextColor(nBackground)
	If bOutline Then
		' Outline: 8 copies of the text, in the opposite color, offset all
		' around (about 1/14 of the text height), under the text. In the
		' copies, a text without outline is replaced by a space of the same
		' font and size: line height and centering unchanged (a 100%
		' transparent text is badly drawn by LibreOffice 25.2)
		For k = 0 To n - 1
			aCopies(k) = IIf(aOutlines(k), aTexts(k), " ")
		Next k
		ReDim Preserve aCopies(n - 1) As String
		dGap = IIf(dValueSize > dLabelSize, dValueSize, dLabelSize) * 2.5
		aDx = Array(-1, 0, 1, -1, 1, -1, 0, 1)
		aDy = Array(-1, -1, -1, 0, 0, 1, 1, 1)
		For k = 0 To 7
			oText = ZoneTextShape(oDoc, oContainer, oZone.Name, aCopies, aSizes, aFonts, aStyles, _
				IIf(nColor = 0, RGB(255, 255, 255), 0), dCx + aDx(k) * dGap, dCy + aDy(k) * dGap)
		Next k
	End If
	oText = ZoneTextShape(oDoc, oContainer, oZone.Name, aTexts, aSizes, aFonts, aStyles, nColor, dCx, dCy)
	Exit Sub
Failed:
	WriteLog("  zone text " & oZone.Name & " not created (" & Error$ & ")")
End Sub

' Center of the zone texts: center of the largest part of the zone (a
' country with islands or overseas territories, a continent), else of the
' whole shape
Sub ZoneCenter(oZone As Object, dCx As Double, dCy As Double)
	Dim aPolys As Variant, aPts As Variant, i As Long, k As Long, n As Long
	Dim dArea As Double, dBest As Double, x0 As Double, x1 As Double, y0 As Double, y1 As Double
	Dim xA As Double, yA As Double, xB As Double, yB As Double
	dCx = oZone.Position.X + oZone.Size.Width / 2
	dCy = oZone.Position.Y + oZone.Size.Height / 2
	On Error GoTo Done
	Select Case oZone.getShapeType()
		Case "com.sun.star.drawing.PolyPolygonShape" : aPolys = oZone.PolyPolygon
		Case "com.sun.star.drawing.ClosedBezierShape" : aPolys = oZone.PolyPolygonBezier.Coordinates
		Case Else : Exit Sub
	End Select
	If UBound(aPolys) < 1 Then Exit Sub
	For i = 0 To UBound(aPolys)
		aPts = aPolys(i)
		n = UBound(aPts)
		dArea = 0
		For k = 0 To n
			xA = aPts(k).X : yA = aPts(k).Y
			xB = aPts((k + 1) Mod (n + 1)).X : yB = aPts((k + 1) Mod (n + 1)).Y
			dArea = dArea + xA * yB - xB * yA
		Next k
		If Abs(dArea) > dBest Then
			dBest = Abs(dArea)
			x0 = aPts(0).X : x1 = x0 : y0 = aPts(0).Y : y1 = y0
			For k = 1 To n
				If aPts(k).X < x0 Then x0 = aPts(k).X
				If aPts(k).X > x1 Then x1 = aPts(k).X
				If aPts(k).Y < y0 Then y0 = aPts(k).Y
				If aPts(k).Y > y1 Then y1 = aPts(k).Y
			Next k
		End If
	Next i
	' Points outside the frame of the shape (unexpected): center of the shape
	If dBest = 0 Or x0 < oZone.Position.X - 1 Or y0 < oZone.Position.Y - 1 _
		Or x1 > oZone.Position.X + oZone.Size.Width + 1 Or y1 > oZone.Position.Y + oZone.Size.Height + 1 Then Exit Sub
	dCx = (x0 + x1) / 2
	dCy = (y0 + y1) / 2
Done:
End Sub

' Text shape (one centered paragraph per text) centered on (dCx, dCy),
' added to the container after the existing shapes
Function ZoneTextShape(oDoc As Object, oContainer As Object, sCode As String, aTexts As Variant, _
		aSizes As Variant, aFonts As Variant, aStyles As Variant, nColor As Long, _
		ByVal dCx As Double, ByVal dCy As Double) As Object
	Dim oText As Object, oEnum As Object, oPara As Object, k As Integer
	oText = oDoc.createInstance("com.sun.star.drawing.TextShape")
	oContainer.add(oText)
	oText.Name = TEXT_PREFIX & " " & sCode
	oText.TextAutoGrowWidth = True
	oText.TextAutoGrowHeight = True
	oText.TextWordWrap = False
	oText.TextLeftDistance = 0
	oText.TextRightDistance = 0
	oText.TextUpperDistance = 0
	oText.TextLowerDistance = 0
	oText.setString(Join(aTexts, Chr(10)))
	oText.CharColor = nColor
	oEnum = oText.getText().createEnumeration()
	Do While oEnum.hasMoreElements() And k <= UBound(aTexts)
		oPara = oEnum.nextElement()
		oPara.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER
		ApplyFont(oPara, aFonts(k), aSizes(k), aStyles(k))
		k = k + 1
	Loop
	CenterShape(oText, dCx, dCy)
	ZoneTextShape = oText
End Function

' Places a shape so that its center is at (dCx, dCy)
Sub CenterShape(oShp As Object, ByVal dCx As Double, ByVal dCy As Double)
	Dim oPos As New com.sun.star.awt.Point
	oPos.X = dCx - oShp.Size.Width / 2
	oPos.Y = dCy - oShp.Size.Height / 2
	oShp.setPosition(oPos)
End Sub

' Adds the value to the existing alternative text:
' "22 - Côtes-d'Armor" -> "22 - Côtes-d'Armor - Value: 666".
' A value added by a previous run is replaced.
Function AltText(sExisting As String, sValue As String) As String
	Dim sPrefix As String, sBase As String, p As Long
	sBase = sExisting
	sPrefix = Tr("alt.value", "")
	If sPrefix <> "" Then
		If Left(sBase, Len(sPrefix)) = sPrefix Then
			sBase = ""
		Else
			p = InStr(1, sBase, " - " & sPrefix, 0)
			If p > 0 Then sBase = Left(sBase, p - 1)
		End If
	End If
	If sBase = "" Then
		AltText = sValue
	Else
		AltText = sBase & " - " & sValue
	End If
End Function

' ---------------------------------------------------------------------------
' Context of a coloring (see T_*): the number format of the first
' numeric value of the values column, to show them as in the
' sheet (decimals, %, thousands…), the options, the document where the
' legend is created (Nothing: no legend) and its title (column header).
' ---------------------------------------------------------------------------
Function TextContext(oDocCalc As Object, oRange As Object, aData As Variant, _
		aOptions As Variant, oLegendDoc As Object) As Variant
	Dim oFormatter As Object, nKey As Long, i As Long, iVal As Integer, sTitle As String, d As Double
	iVal = aOptions(O_COL_VALUE)
	On Error GoTo Failed
	oFormatter = CreateUnoService("com.sun.star.util.NumberFormatter")
	oFormatter.attachNumberFormatsSupplier(oDocCalc)
	nKey = 0
	For i = 0 To UBound(aData)
		If UBound(aData(i)) >= iVal Then
			Select Case VarType(aData(i)(iVal))
				Case 2, 3, 4, 5, 6, 7
					nKey = oRange.getCellByPosition(iVal, i).NumberFormat
					Exit For
			End Select
		End If
	Next i
	' Header: 1st row of the values column, if it is text
	If UBound(aData) >= 0 Then
		If UBound(aData(0)) >= iVal Then
			If VarType(aData(0)(iVal)) = 8 Then
				If Not ToNumber(aData(0)(iVal), d) Then sTitle = Trim(aData(0)(iVal))
			End If
		End If
	End If
	TextContext = Array(oFormatter, nKey, aOptions, oLegendDoc, sTitle)
	Exit Function
Failed:
	WriteLog("Value format not read (" & Error$ & "): standard format")
	TextContext = Array(Nothing, 0, aOptions, oLegendDoc, sTitle)
End Function

' Text options (see O_*), sizes brought within the limits
Function DefaultOptions() As Variant
	DefaultOptions = Array(False, TEXT_SIZE, False, TEXT_SIZE, TEXT_FONT, "", TEXT_FONT, "", _
		0, 1, 2, False, True, "", OUTLINE_ON, OUTLINE_ON)
End Function

' Option kept in the preferences (neither columns nor classes, specific to
' each table)
Function IsRememberedOption(k As Integer) As Boolean
	IsRememberedOption = (InStr("ck", Mid(OPTION_TYPES, k + 1, 1)) = 0)
End Function

Function OptionName(k As Integer) As String
	OptionName = Split(OPTION_NAMES, "|")(k)
End Function

' Option k read from a text (link, preferences)
Function OptionFromText(k As Integer, s As String) As Variant
	Dim aDefault As Variant
	Select Case Mid(OPTION_TYPES, k + 1, 1)
		Case "b"
			OptionFromText = (s = "1")
		Case "n"
			OptionFromText = ValidSize(Val(s))
		Case "c"
			If Val(s) < 0 Or Val(s) >= MAX_COLUMNS Or s = "" Then
				aDefault = DefaultOptions()
				OptionFromText = aDefault(k)
			Else
				OptionFromText = CInt(Val(s))
			End If
		Case Else
			OptionFromText = s
	End Select
End Function

' Option k written as text (without |, the link separator)
Function OptionToText(k As Integer, v As Variant) As String
	If Mid(OPTION_TYPES, k + 1, 1) = "b" Then
		OptionToText = IIf(v, "1", "0")
	Else
		OptionToText = Replace(CStr(v), "|", "")
	End If
End Function

Function ValueText(vText As Variant, ByVal dVal As Double) As String
	On Error GoTo Failed
	ValueText = vText(T_FORMATTER).convertNumberToString(vText(T_FORMAT), dVal)
	Exit Function
Failed:
	ValueText = CStr(dVal)
End Function

' Size of the values brought between TEXT_SIZE_MIN and TEXT_SIZE_MAX
' (TEXT_SIZE if missing or unreadable)
Function ValidSize(v As Variant) As Integer
	Dim d As Double
	On Error GoTo UseDefault
	d = CDbl(v)
	If d <= 0 Then GoTo UseDefault
	If d < TEXT_SIZE_MIN Then d = TEXT_SIZE_MIN
	If d > TEXT_SIZE_MAX Then d = TEXT_SIZE_MAX
	ValidSize = CInt(d)
	Exit Function
UseDefault:
	ValidSize = TEXT_SIZE
End Function

' Width (1/100 mm) of a draw page, a shape or a collection
Function ZonesWidth(o As Object) As Double
	Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
	On Error GoTo UseDefault
	If o.supportsService("com.sun.star.drawing.DrawPage") Then
		ZonesWidth = o.Width
		Exit Function
	End If
	ZonesExtent(o, x1, y1, x2, y2)
	If x2 > x1 Then ZonesWidth = x2 - x1 Else ZonesWidth = CALC_MAP_WIDTH
	Exit Function
UseDefault:
	ZonesWidth = CALC_MAP_WIDTH
End Function

' Rectangle (1/100 mm) taken by the zones of a container, texts and legend
' excluded
Sub ZonesExtent(o As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
	Dim i As Long, oZ As Object, bInit As Boolean, p As Variant, t As Variant
	For i = 0 To o.getCount() - 1
		oZ = o.getByIndex(i)
		If Not IsAddOn(oZ.Name) Then
			p = oZ.Position
			t = oZ.Size
			If Not bInit Then
				x1 = p.X : y1 = p.Y : x2 = p.X + t.Width : y2 = p.Y + t.Height
				bInit = True
			Else
				If p.X < x1 Then x1 = p.X
				If p.Y < y1 Then y1 = p.Y
				If p.X + t.Width > x2 Then x2 = p.X + t.Width
				If p.Y + t.Height > y2 Then y2 = p.Y + t.Height
			End If
		End If
	Next i
End Sub

' Zone text: one paragraph per text (label, value), each with
' its size, font and style, centered; the block is centered in the
' zone, without line wrapping and without enlarging the shape
Sub WriteParagraphs(oShape As Object, aTexts As Variant, aSizes As Variant, aFonts As Variant, _
		aStyles As Variant, nColor As Long)
	Dim oEnum As Object, oPara As Object, k As Integer
	oShape.setString(Join(aTexts, Chr(10)))
	oShape.TextAutoGrowHeight = False
	oShape.TextAutoGrowWidth = False
	oShape.TextWordWrap = False
	oShape.TextHorizontalAdjust = com.sun.star.drawing.TextHorizontalAdjust.CENTER
	oShape.TextVerticalAdjust = com.sun.star.drawing.TextVerticalAdjust.CENTER
	oShape.CharColor = nColor
	oEnum = oShape.getText().createEnumeration()
	k = 0
	Do While oEnum.hasMoreElements() And k <= UBound(aSizes)
		oPara = oEnum.nextElement()
		oPara.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER
		ApplyFont(oPara, aFonts(k), aSizes(k), aStyles(k))
		k = k + 1
	Loop
End Sub

' Font, size and style (letters B, I, U) of a text or a paragraph
Sub ApplyFont(o As Object, sFont As String, dSize As Double, sStyle As String)
	If sFont <> "" Then o.CharFontName = sFont
	o.CharHeight = dSize
	If InStr(sStyle, "B") > 0 Then
		o.CharWeight = com.sun.star.awt.FontWeight.BOLD
	Else
		o.CharWeight = com.sun.star.awt.FontWeight.NORMAL
	End If
	If InStr(sStyle, "I") > 0 Then
		o.CharPosture = com.sun.star.awt.FontSlant.ITALIC
	Else
		o.CharPosture = com.sun.star.awt.FontSlant.NONE
	End If
	If InStr(sStyle, "U") > 0 Then
		o.CharUnderline = com.sun.star.awt.FontUnderline.SINGLE
	Else
		o.CharUnderline = com.sun.star.awt.FontUnderline.NONE
	End If
End Sub

' Style (letters B, I, U) of a text or a paragraph
Function StyleOf(o As Object) As String
	Dim sStyle As String
	If o.CharWeight > com.sun.star.awt.FontWeight.NORMAL Then sStyle = sStyle & "B"
	If o.CharPosture <> com.sun.star.awt.FontSlant.NONE Then sStyle = sStyle & "I"
	If o.CharUnderline <> com.sun.star.awt.FontUnderline.NONE Then sStyle = sStyle & "U"
	StyleOf = sStyle
End Function

' Copies the text of a zone (paragraphs, fonts, styles and sizes,
' multiplied by dScale)
Sub CopyText(oSource As Object, oCopy As Object, dScale As Double)
	Dim oEnum As Object, oPara As Object, n As Integer
	Dim aTexts() As String, aSizes() As Double, aFonts() As String, aStyles() As String
	oEnum = oSource.getText().createEnumeration()
	Do While oEnum.hasMoreElements()
		oPara = oEnum.nextElement()
		ReDim Preserve aTexts(n) As String
		ReDim Preserve aSizes(n) As Double
		ReDim Preserve aFonts(n) As String
		ReDim Preserve aStyles(n) As String
		aTexts(n) = oPara.getString()
		aSizes(n) = oPara.CharHeight * dScale
		aFonts(n) = oPara.CharFontName
		aStyles(n) = StyleOf(oPara)
		n = n + 1
	Loop
	If n > 0 Then WriteParagraphs(oCopy, aTexts, aSizes, aFonts, aStyles, oSource.CharColor)
End Sub

' ---------------------------------------------------------------------------
' Custom classes: bounds b0 < b1 < … < bn and one color per class
' [b(k), b(k+1)[ (the last one includes bn). Text (option O_CLASSES, link):
' "b0;b1;…;bn/rrggbb;…" (bounds in Str format, decimal point).
' ---------------------------------------------------------------------------
Function ClassesText(aBounds As Variant, aColors As Variant, n As Integer) As String
	Dim k As Integer, s As String, c As String
	For k = 0 To n
		s = s & IIf(k > 0, ";", "") & Trim(Str(aBounds(k)))
	Next k
	s = s & "/"
	For k = 0 To n - 1
		c = Hex(aColors(k))
		s = s & IIf(k > 0, ";", "") & String(6 - Len(c), "0") & c
	Next k
	ClassesText = LCase(s)
End Function

' Reads a classes text; False if it is empty or malformed
Function ReadClasses(sText As String, aBounds As Variant, aColors As Variant) As Boolean
	Dim aParts As Variant, aB As Variant, aC As Variant, k As Integer, n As Integer
	On Error GoTo Failed
	ReadClasses = False
	If sText = "" Then Exit Function
	aParts = Split(sText, "/")
	If UBound(aParts) <> 1 Then Exit Function
	aB = Split(aParts(0), ";")
	aC = Split(aParts(1), ";")
	n = UBound(aC) + 1
	If n < 1 Or UBound(aB) <> n Then Exit Function
	ReDim aBounds(n)
	ReDim aColors(n - 1)
	For k = 0 To n
		aBounds(k) = Val(aB(k))
		If k > 0 Then
			If aBounds(k) <= aBounds(k - 1) Then Exit Function
		End If
	Next k
	For k = 0 To n - 1
		aColors(k) = HexColor(CStr(aC(k)))
		If aColors(k) < 0 Then Exit Function
	Next k
	ReadClasses = True
	Exit Function
Failed:
	ReadClasses = False
End Function

' Class of a value (-1 outside the classes)
Function ClassOf(aBounds As Variant, ByVal dVal As Double) As Integer
	Dim k As Integer, n As Integer
	n = UBound(aBounds)
	ClassOf = -1
	If dVal < aBounds(0) Or dVal > aBounds(n) Then Exit Function
	For k = 0 To n - 1
		If dVal < aBounds(k + 1) Or k = n - 1 Then
			ClassOf = k
			Exit Function
		End If
	Next k
End Function

' Linear distribution in nWanted classes (0: one per palette color,
' 5 for a continuous gradient); palette colors, stretched into a gradient if
' the number of classes differs
Sub LinearClasses(aPalette As Variant, ByVal dMin As Double, ByVal dMax As Double, ByVal nWanted As Integer, _
		aBounds As Variant, aColors As Variant, n As Integer)
	Dim k As Integer, aPal As Variant
	aPal = aPalette(P_COLORS)
	If nWanted > 0 Then
		n = nWanted
	ElseIf aPalette(P_CONTINUOUS) Then
		n = 5
	Else
		n = UBound(aPal) + 1
	End If
	LinearBounds(dMin, dMax, n, aBounds)
	ReDim aColors(n - 1)
	For k = 0 To n - 1
		If Not aPalette(P_CONTINUOUS) And n = UBound(aPal) + 1 Then
			aColors(k) = aPal(k)
		ElseIf n = 1 Then
			aColors(k) = GradientColor(aPal, 0.5)
		Else
			aColors(k) = GradientColor(aPal, k / (n - 1))
		End If
	Next k
End Sub

' n equal intervals between dMin and dMax, rounded bounds (outwards
' at the ends)
Sub LinearBounds(ByVal dMin As Double, ByVal dMax As Double, n As Integer, aBounds As Variant)
	Dim k As Integer, dStep As Double
	ReDim aBounds(n)
	dStep = RoundingStep(dMax - dMin)
	For k = 0 To n
		aBounds(k) = Int((dMin + (dMax - dMin) * k / n) / dStep + 0.5) * dStep
	Next k
	aBounds(0) = Int(dMin / dStep) * dStep
	aBounds(n) = -Int(-dMax / dStep) * dStep
	For k = 1 To n
		If aBounds(k) <= aBounds(k - 1) Then aBounds(k) = aBounds(k - 1) + dStep
	Next k
End Sub

' Rounding step of the bounds: a power of 10, about 1/100 of the range
Function RoundingStep(ByVal dRange As Double) As Double
	If dRange <= 0 Then
		RoundingStep = 1
	Else
		RoundingStep = 10 ^ (Int(Log(dRange) / Log(10)) - 2)
	End If
End Function

' ---------------------------------------------------------------------------
' Legend: one box per class (5 shades for a continuous gradient), with its
' bounds if requested, and "No data"; to the right of the zones, in the same
' container (group of the inserted map, or draw page). All its shapes
' are named LEGEND_PREFIX.
' ---------------------------------------------------------------------------
Function IsLegend(sName As String) As Boolean
	IsLegend = (Left(sName, Len(LEGEND_PREFIX)) = LEGEND_PREFIX) _
		Or (Left(sName, Len(OLD_LEGEND_PREFIX)) = OLD_LEGEND_PREFIX)
End Function

' Outline option set: any value other than "none" (including the former
' "etiquette"), or "aucune" written by versions up to 0.19
Function HasOutline(sValue As String) As Boolean
	HasOutline = (sValue <> OUTLINE_NONE And sValue <> OLD_OUTLINE_NONE)
End Function

' Shape added by CalcMap to the map (zone text, legend): not a zone
Function IsAddOn(sName As String) As Boolean
	IsAddOn = IsLegend(sName) Or (Left(sName, Len(TEXT_PREFIX)) = TEXT_PREFIX)
End Function

Function AcceptsLegend(oContainer As Object) As Boolean
	On Error GoTo Failed
	AcceptsLegend = oContainer.supportsService("com.sun.star.drawing.GroupShape") _
		Or oContainer.supportsService("com.sun.star.drawing.DrawPage")
	Exit Function
Failed:
	AcceptsLegend = False
End Function

' Removes zone texts and legend from a container (before recreating them)
Sub RemoveAddOns(oContainer As Object)
	Dim i As Long
	On Error GoTo Failed
	If Not AcceptsLegend(oContainer) Then Exit Sub
	For i = oContainer.getCount() - 1 To 0 Step -1
		If IsAddOn(oContainer.getByIndex(i).Name) Then oContainer.remove(oContainer.getByIndex(i))
	Next i
	Exit Sub
Failed:
	WriteLog("ERROR removing the texts and the legend: " & Error$)
End Sub

Sub CreateLegend(vText As Variant, oContainer As Object, aPalette As Variant, dMin As Double, _
		dMax As Double, bHasData As Boolean, bWithout As Boolean, dScale As Double)
	Dim oDoc As Object, aOptions As Variant, aColors As Variant, bContinuous As Boolean
	Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x As Double, y As Double, xMax As Double
	Dim dSize As Double, dBox As Double, dStep As Double, k As Integer, n As Integer
	Dim nColor As Long, dPos As Double, sText As String, bValues As Boolean
	Dim bCustom As Boolean, aBounds As Variant, aCustomColors As Variant
	On Error GoTo Failed
	If Not AcceptsLegend(oContainer) Then
		WriteLog("No legend possible for this map")
		Exit Sub
	End If
	oDoc = vText(T_DOCUMENT)
	aOptions = vText(T_OPTIONS)
	bValues = aOptions(O_LEGEND_VALUES)
	aColors = aPalette(P_COLORS)
	bContinuous = aPalette(P_CONTINUOUS)
	ZonesExtent(oContainer, x1, y1, x2, y2)
	' Text in the font and size of the values (without their style, title in
	' bold); box a little higher
	' than the text (1 pt = 35.28 hundredths of a mm)
	dSize = aOptions(O_SIZE_VALUES) * dScale
	dBox = dSize * 35.28 * 1.3
	dStep = dBox * 1.3
	x = x2 + dBox
	y = y1
	xMax = x
	If vText(T_TITLE) <> "" And bValues Then
		xMax = LegendText(oDoc, oContainer, vText(T_TITLE), x, y, dBox, dSize, _
			aOptions(O_FONT_VALUES), "B", xMax)
		y = y + dStep
	End If
	bCustom = ReadClasses(aOptions(O_CLASSES), aBounds, aCustomColors)
	If bHasData Then
		If bCustom Then
			n = UBound(aCustomColors) + 1
		ElseIf bContinuous Then
			n = 5
		Else
			n = UBound(aColors) + 1
		End If
		For k = 0 To n - 1
			If bCustom Then
				nColor = aCustomColors(k)
				sText = ValueText(vText, aBounds(k)) & " – " & ValueText(vText, aBounds(k + 1))
			ElseIf bContinuous Then
				dPos = k / (n - 1)
				nColor = GradientColor(aColors, dPos)
				sText = ValueText(vText, dMin + dPos * (dMax - dMin))
			Else
				nColor = aColors(k)
				sText = ValueText(vText, dMin + (dMax - dMin) * k / n) & " – " _
					& ValueText(vText, dMin + (dMax - dMin) * (k + 1) / n)
			End If
			LegendBox(oDoc, oContainer, nColor, x, y, dBox * 1.5, dBox)
			If bValues Then xMax = LegendText(oDoc, oContainer, sText, x + dBox * 1.9, y, dBox, _
				dSize, aOptions(O_FONT_VALUES), "", xMax)
			y = y + dStep
		Next k
	End If
	If bWithout Then
		LegendBox(oDoc, oContainer, aPalette(P_EMPTY), x, y, dBox * 1.5, dBox)
		If bValues Then xMax = LegendText(oDoc, oContainer, Tr("legend.no_data"), x + dBox * 1.9, y, _
			dBox, dSize, aOptions(O_FONT_VALUES), "", xMax)
	End If
	' Draw page: widened if the legend goes beyond it
	If oContainer.supportsService("com.sun.star.drawing.DrawPage") Then
		If xMax + dBox > oContainer.Width Then oContainer.Width = xMax + dBox
	End If
	Exit Sub
Failed:
	WriteLog("ERROR creating the legend " & Err & " (line " & Erl & "): " & Error$)
End Sub

Sub LegendBox(oDoc As Object, oContainer As Object, nColor As Long, x As Double, y As Double, _
		dWidth As Double, dHeight As Double)
	Dim oBox As Object
	Dim oPos As New com.sun.star.awt.Point, oSize As New com.sun.star.awt.Size
	oBox = oDoc.createInstance("com.sun.star.drawing.RectangleShape")
	oContainer.add(oBox)
	oBox.Name = LEGEND_PREFIX
	oPos.X = x : oPos.Y = y
	oSize.Width = dWidth : oSize.Height = dHeight
	oBox.setPosition(oPos)
	oBox.setSize(oSize)
	oBox.FillStyle = com.sun.star.drawing.FillStyle.SOLID
	oBox.FillColor = nColor
	oBox.LineColor = RGB(96, 96, 96)
End Sub

' Legend text, vertically centered on a box of height dHeight;
' returns the rightmost right edge of the legend
Function LegendText(oDoc As Object, oContainer As Object, sText As String, x As Double, y As Double, _
		dHeight As Double, dSize As Double, sFont As String, sStyle As String, xMax As Double) As Double
	Dim oText As Object
	Dim oPos As New com.sun.star.awt.Point, oSize As New com.sun.star.awt.Size
	oText = oDoc.createInstance("com.sun.star.drawing.TextShape")
	oContainer.add(oText)
	oText.Name = LEGEND_PREFIX
	oText.TextAutoGrowWidth = True
	oText.TextAutoGrowHeight = False
	oText.TextWordWrap = False
	oText.TextVerticalAdjust = com.sun.star.drawing.TextVerticalAdjust.CENTER
	' Anchored left: the text grows to the right (otherwise on both sides)
	oText.TextHorizontalAdjust = com.sun.star.drawing.TextHorizontalAdjust.LEFT
	oText.TextLeftDistance = 0
	oText.TextRightDistance = 0
	oText.setString(sText)
	oText.CharColor = RGB(0, 0, 0)
	ApplyFont(oText, sFont, dSize, sStyle)
	oPos.X = x : oPos.Y = y
	oSize.Width = dHeight : oSize.Height = dHeight
	oText.setSize(oSize)
	oText.setPosition(oPos)
	LegendText = xMax
	If x + oText.Size.Width > xMax Then LegendText = x + oText.Size.Width
End Function

' Black on a light background, white on a dark one
Function ContrastTextColor(nBackground As Long) As Long
	If 0.299 * Red(nBackground) + 0.587 * Green(nBackground) + 0.114 * Blue(nBackground) < 140 Then
		ContrastTextColor = RGB(255, 255, 255)
	Else
		ContrastTextColor = RGB(0, 0, 0)
	End If
End Function

' ---------------------------------------------------------------------------
' Data reading: fills oMap (code -> value) and oLabels (code ->
' label) from columns iCode, iVal, iLab. Returns the number of
' rows kept. Each discarded row is traced in the log.
' ---------------------------------------------------------------------------
Function ReadData(aData As Variant, oMap As Object, dMin As Double, dMax As Double, _
		oLabels As Object, iCode As Integer, iVal As Integer, iLab As Integer) As Long
	Dim i As Long, sCode As String, vVal As Variant, dVal As Double, vLab As Variant
	Dim n As Long, nRow As Long, bInit As Boolean
	On Error GoTo Failed
	ReadData = 0
	If Not IsArray(aData) Then Exit Function

	For i = LBound(aData) To UBound(aData)
		nRow = i - LBound(aData) + 1
		sCode = "" : vVal = Empty
		If IsArray(aData(i)) Then
			If UBound(aData(i)) >= iCode Then sCode = NormalizeCode(aData(i)(iCode))
			If UBound(aData(i)) >= iVal Then vVal = aData(i)(iVal)
			' Label, kept even without a value
			If sCode <> "" And UBound(aData(i)) >= iLab Then
				vLab = aData(i)(iLab)
				If Not IsBlank(vLab) Then oLabels.put(sCode, Trim(CStr(vLab)))
			End If
		End If

		If sCode = "" Then
			If Not IsBlank(vVal) Then WriteLog("Row " & nRow & " skipped: empty code")
		ElseIf IsBlank(vVal) Then
			WriteLog("Row " & nRow & " (" & sCode & "): empty value -> default color")
		ElseIf Not ToNumber(vVal, dVal) Then
			WriteLog("Row " & nRow & " (" & sCode & "): non-numeric value '" & CStr(vVal) _
				& "' -> skipped (header?)")
		Else
			If oMap.containsKey(sCode) Then
				WriteLog("Row " & nRow & " (" & sCode & "): duplicate code, the last value is kept")
			Else
				n = n + 1
			End If
			oMap.put(sCode, CreateUnoValue("double", dVal))
			If Not bInit Then
				dMin = dVal : dMax = dVal : bInit = True
			Else
				If dVal < dMin Then dMin = dVal
				If dVal > dMax Then dMax = dVal
			End If
			ReadData = n
		End If
	Next i
	WriteLog("Rows kept: " & n & " code(s)")
	Exit Function
Failed:
	WriteLog("ERROR reading the data, row " & nRow & ": " & Error$)
	ReadData = n
End Function

' Code -> normalized text ("" if unusable)
Function NormalizeCode(v As Variant) As String
	Dim s As String
	On Error GoTo Failed
	NormalizeCode = ""
	Select Case VarType(v)
		Case 2, 3, 4, 5, 6, 7           ' numbers (e.g. 75, 1001)
			If v >= 0 And v = Int(v) Then NormalizeCode = Format(v, "0")
		Case 8                           ' text (e.g. "01", "2A")
			s = Trim(Replace(v, Chr(160), " "))
			NormalizeCode = UCase(s)
	End Select
	Exit Function
Failed:
	NormalizeCode = ""
End Function

Function IsBlank(v As Variant) As Boolean
	On Error GoTo Failed
	IsBlank = True
	If IsEmpty(v) Or IsNull(v) Then Exit Function
	If VarType(v) = 8 Then
		If Trim(Replace(v, Chr(160), " ")) = "" Then Exit Function
	End If
	IsBlank = False
	Exit Function
Failed:
	IsBlank = True
End Function

' Value -> number. Accepts numbers and texts like "1 234,5",
' "-12.3", "+7". Returns False for anything else (headers, "n.d.", errors).
Function ToNumber(v As Variant, dVal As Double) As Boolean
	Dim s As String, c As String, k As Integer, nDigits As Integer, nSep As Integer
	On Error GoTo Failed
	ToNumber = False
	Select Case VarType(v)
		Case 2, 3, 4, 5, 6, 7
			dVal = v
			ToNumber = True
			Exit Function
		Case 8
			s = Replace(v, Chr(160), "")
			s = Replace(s, Chr(8239), "")    ' narrow no-break space
			s = Replace(s, " ", "")
			s = Replace(s, ",", ".")
			If s = "" Then Exit Function
			For k = 1 To Len(s)
				c = Mid(s, k, 1)
				If c >= "0" And c <= "9" Then
					nDigits = nDigits + 1
				ElseIf c = "." Then
					nSep = nSep + 1
					If nSep > 1 Then Exit Function
				ElseIf (c = "-" Or c = "+") And k = 1 Then
					' sign accepted at the start
				Else
					Exit Function
				End If
			Next k
			If nDigits = 0 Then Exit Function
			dVal = Val(s)
			ToNumber = True
	End Select
	Exit Function
Failed:
	ToNumber = False
End Function

' Looks up the value of a code; tolerates leading zeros lost by Calc
' (e.g. department "01" typed as the number 1, municipality "01001" as 1001)
Function FindValue(oMap As Object, sName As String, dVal As Double, sFoundKey As String) As Boolean
	On Error GoTo Failed
	FindValue = False
	sFoundKey = FoundKey(oMap, sName)
	If sFoundKey = "" Then Exit Function
	dVal = oMap.get(sFoundKey)
	FindValue = True
	Exit Function
Failed:
	FindValue = False
End Function

' Label of a zone ("" if missing)
Function FindLabel(oLabels As Object, sName As String) As String
	Dim sKey As String
	On Error GoTo Failed
	sKey = FoundKey(oLabels, sName)
	If sKey <> "" Then FindLabel = oLabels.get(sKey)
	Exit Function
Failed:
	FindLabel = ""
End Function

' Key of oMap matching the name of a zone, tolerating lost leading
' zeros ("" if missing)
Function FoundKey(oMap As Object, sName As String) As String
	Dim sKey As String
	FoundKey = ""
	sKey = UCase(Trim(sName))
	Do While sKey <> ""
		If oMap.containsKey(sKey) Then
			FoundKey = sKey
			Exit Function
		End If
		If Left(sKey, 1) <> "0" Then Exit Do
		sKey = Mid(sKey, 2)
	Loop
End Function

' ---------------------------------------------------------------------------
' Selection: SelectedRange returns the first 2 columns of the
' selection as they are (whole columns included); DataRange
' limits them to the used area of the sheet. Null if unusable.
' ---------------------------------------------------------------------------
Function SelectedRange(oSel As Object) As Object
	Dim oRange As Object, a As Variant, nLast As Long
	On Error GoTo Failed
	SelectedRange = Nothing
	If IsNull(oSel) Then Exit Function
	If oSel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
		If oSel.getCount() = 0 Then Exit Function
		oRange = oSel.getByIndex(0)
		WriteLog("Multiple selection: only the first range is used.")
	ElseIf oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then
		oRange = oSel
	Else
		Exit Function
	End If
	a = oRange.getRangeAddress()
	If a.EndColumn - a.StartColumn < 1 Then Exit Function
	' At least 3 columns (code, value, label: the 3rd is taken even outside
	' the selection), at most MAX_COLUMNS
	nLast = a.EndColumn
	If nLast < a.StartColumn + 2 Then nLast = a.StartColumn + 2
	If nLast > a.StartColumn + MAX_COLUMNS - 1 Then nLast = a.StartColumn + MAX_COLUMNS - 1
	SelectedRange = oRange.getSpreadsheet().getCellRangeByPosition(a.StartColumn, a.StartRow, _
		nLast, a.EndRow)
	Exit Function
Failed:
	WriteLog("ERROR on the selection: " & Error$)
	SelectedRange = Nothing
End Function

Function DataRange(oRange As Object) As Object
	Dim oSheet As Object, oCursor As Object
	Dim a As Variant, nLast As Long
	On Error GoTo Failed
	DataRange = Nothing
	If IsNull(oRange) Then Exit Function
	a = oRange.getRangeAddress()

	oSheet = oRange.getSpreadsheet()
	oCursor = oSheet.createCursor()
	oCursor.gotoEndOfUsedArea(False)
	nLast = oCursor.getRangeAddress().EndRow
	If a.EndRow < nLast Then nLast = a.EndRow
	If a.StartRow > nLast Then Exit Function

	If a.EndColumn < a.StartColumn + 2 Then a.EndColumn = a.StartColumn + 2   ' links of versions 0.1 to 0.3
	DataRange = oSheet.getCellRangeByPosition(a.StartColumn, a.StartRow, a.EndColumn, nLast)
	Exit Function
Failed:
	WriteLog("ERROR on the selection: " & Error$)
	DataRange = Nothing
End Function

' The 1st row of a range is a header if a column has text there
' above a number (e.g. "Value" above 389). Codes or names alone
' are not enough: they are text in the data too.
Function IsHeaderRow(oRange As Object) As Boolean
	Dim c As Long
	On Error GoTo Failed
	If oRange.getRows().getCount() < 2 Then Exit Function
	For c = 0 To oRange.getColumns().getCount() - 1
		If oRange.getCellByPosition(c, 0).getType() = com.sun.star.table.CellContentType.TEXT _
				And oRange.getCellByPosition(c, 1).getType() = com.sun.star.table.CellContentType.VALUE Then
			IsHeaderRow = True
			Exit Function
		End If
	Next c
Failed:
End Function

' Column names of a range for the dialog: letter, header (1st row)
' if it is text, and first values as displayed, e.g.
' "B - Value: 531, 523, 840"
Function ColumnNames(oRange As Object) As Variant
	Dim aNames() As String, c As Long, n As Long, sHeader As String
	Dim nRows As Long, nStart As Long, r As Long, sValues As String, nValues As Integer, sCell As String
	On Error GoTo Failed
	n = oRange.getColumns().getCount()
	nRows = oRange.getRows().getCount()
	ReDim aNames(n - 1) As String
	If IsHeaderRow(oRange) Then nStart = 1
	For c = 0 To n - 1
		aNames(c) = oRange.getColumns().getByIndex(c).getName()
		If nStart = 1 Then
			sHeader = Trim(oRange.getCellByPosition(c, 0).getString())
			If Len(sHeader) > 30 Then sHeader = Left(sHeader, 29) & "…"
			If sHeader <> "" Then aNames(c) = aNames(c) & " - " & sHeader
		End If
		' First non-empty values (among the next PREVIEW_ROWS rows)
		sValues = ""
		nValues = 0
		For r = nStart To nRows - 1
			If r >= nStart + PREVIEW_ROWS Or nValues = PREVIEW_VALUES Then Exit For
			sCell = Trim(oRange.getCellByPosition(c, r).getString())
			If sCell <> "" Then
				If Len(sCell) > 15 Then sCell = Left(sCell, 14) & "…"
				sValues = sValues & IIf(nValues > 0, ", ", "") & sCell
				nValues = nValues + 1
			End If
		Next r
		' More values follow: ", ..."
		If nValues = PREVIEW_VALUES Then
			Do While r < nRows And r < nStart + PREVIEW_ROWS
				If Trim(oRange.getCellByPosition(c, r).getString()) <> "" Then
					sValues = sValues & ", ..."
					Exit Do
				End If
				r = r + 1
			Loop
		End If
		If sValues <> "" Then aNames(c) = aNames(c) & ": " & sValues
	Next c
	ColumnNames = aNames
	Exit Function
Failed:
	WriteLog("ERROR reading the headers: " & Error$)
	ColumnNames = Array("1", "2", "3")
End Function

Function RangeName(oRange As Object) As String
	On Error GoTo Failed
	RangeName = oRange.AbsoluteName
	Exit Function
Failed:
	RangeName = "?"
End Function

Function PickOdgFile() As String
	Dim oPicker As Object
	On Error GoTo Failed
	PickOdgFile = ""
	oPicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
	oPicker.appendFilter(Tr("picker.filter"), "*.odg")
	oPicker.setTitle(Tr("picker.title"))
	If oPicker.execute() = 1 Then PickOdgFile = oPicker.getFiles()(0)
	Exit Function
Failed:
	WriteLog("ERROR in the file choice: " & Error$)
	PickOdgFile = ""
End Function

' ---------------------------------------------------------------------------
' Choice of the base map (ODG Atlas catalog) and of the palette. Returns
' the URL of the file to open ("" if canceled); bFromCache tells a map from
' the cache; aPalette receives the chosen palette (the default one otherwise);
' bInsert tells whether the map must be inserted into the Calc sheet,
' aOptions receives the options (columns, values, labels, legend);
' aColumns: column names of the selection.
' ---------------------------------------------------------------------------
Function ChooseBaseMap(oDocCalc As Object, bFromCache As Boolean, aPalette As Variant, _
		bInsert As Boolean, aOptions As Variant, aColumns As Variant, aData As Variant) As String
	Dim sFolder As String, aMaps As Variant, bOnline As Boolean, aPalettes As Variant
	Dim sState As String, iChoice As Integer, iPalette As Integer, nRes As Integer
	Dim sPrefMap As String, sPrefPalette As String
	On Error GoTo Failed
	ChooseBaseMap = ""
	bFromCache = False
	aPalettes = LoadPalettes()
	ReadPreferences(sPrefMap, sPrefPalette, bInsert, aOptions)
	aPalette = aPalettes(PaletteIndex(aPalettes, sPrefPalette))
	sFolder = MapsFolder()
	If sFolder = "" Then
		ChooseBaseMap = PickOdgFile()
		Exit Function
	End If
	aMaps = LoadCatalog(sFolder, bOnline)

	If bOnline Then
		sState = Tr("state.online", ConvertFromURL(sFolder))
	ElseIf UBound(aMaps) >= 0 Then
		sState = Tr("state.offline")
	Else
		sState = Tr("state.offline_empty")
	End If

	aDialogData = aData
	nRes = MapDialog(aMaps, aPalettes, sState, sPrefMap, sPrefPalette, _
		iChoice, iPalette, bInsert, aOptions, aColumns, ParentWindow(oDocCalc), False)
	WriteLog("Dialog: result " & nRes & ", map no. " & iChoice & ", palette no. " & iPalette _
		& ", insert into Calc: " & bInsert & ", " & OptionsText(aOptions))
	If (nRes = 1 Or nRes = 3) And iPalette >= 0 Then
		aPalette = aPalettes(iPalette)
		If nRes = 1 And iChoice >= 0 Then
			sPrefMap = aMaps(iChoice)(C_ID)
			If sPrefMap = "" Then sPrefMap = aMaps(iChoice)(C_FILE)
		End If
		WritePreferences(sPrefMap, aPalette(P_ID), bInsert, aOptions)
	End If
	If nRes = 1 And iChoice >= 0 Then
		WriteLog("Chosen map: " & aMaps(iChoice)(C_TITLE))
		ChooseBaseMap = LocalMap(aMaps(iChoice), sFolder, bOnline, oDocCalc)
		bFromCache = (ChooseBaseMap <> "")
	ElseIf nRes = 3 Or nRes < 0 Then
		' 3: "Other file…"; negative: dialog not shown
		ChooseBaseMap = PickOdgFile()
	End If
	Exit Function
Failed:
	WriteLog("ERROR in the map choice " & Err & " (line " & Erl & "): " & Error$)
	ChooseBaseMap = ""
End Function

' Text options, for the log
Function OptionsText(aOptions As Variant) As String
	Dim k As Integer, s As String
	For k = 0 To OPTION_COUNT - 1
		s = s & IIf(k > 0, ", ", "") & OptionName(k) & "=" & OptionToText(k, aOptions(k))
	Next k
	OptionsText = s
End Function

' Last choices (preferences.properties in the cache); insertion into Calc
' enabled by default. Columns, specific to each table, are not
' remembered (1st, 2nd and 3rd by default).
Sub ReadPreferences(sMap As String, sPalette As String, bInsert As Boolean, aOptions As Variant)
	Dim oPrefs As Object, sFolder As String, k As Integer
	On Error GoTo Failed
	sMap = ""
	sPalette = ""
	bInsert = True
	aOptions = DefaultOptions()
	sFolder = MapsFolder()
	If sFolder = "" Then Exit Sub
	oPrefs = com.sun.star.container.EnumerableMap.create("string", "string")
	ReadProperties(sFolder & PREFERENCES_NAME, oPrefs)
	If oPrefs.containsKey("map") Then sMap = oPrefs.get("map")
	If oPrefs.containsKey("palette") Then sPalette = oPrefs.get("palette")
	If oPrefs.containsKey("insert") Then bInsert = (oPrefs.get("insert") <> "0")
	For k = 0 To OPTION_COUNT - 1
		If IsRememberedOption(k) And oPrefs.containsKey(OptionName(k)) Then
			aOptions(k) = OptionFromText(k, oPrefs.get(OptionName(k)))
		End If
	Next k
	Exit Sub
Failed:
	WriteLog("Unreadable preferences: " & Error$)
End Sub

Sub WritePreferences(sMap As String, sPalette As String, bInsert As Boolean, aOptions As Variant)
	Dim sFolder As String, sText As String, k As Integer
	On Error GoTo Failed
	sFolder = MapsFolder()
	If sFolder = "" Then Exit Sub
	sText = "# CalcMap: last choices" & Chr(10) & "map=" & sMap & Chr(10) _
		& "palette=" & sPalette & Chr(10) & "insert=" & IIf(bInsert, "1", "0") & Chr(10)
	For k = 0 To OPTION_COUNT - 1
		If IsRememberedOption(k) Then sText = sText & OptionName(k) & "=" & OptionToText(k, aOptions(k)) & Chr(10)
	Next k
	WriteText(sFolder & PREFERENCES_NAME, sText)
	Exit Sub
Failed:
	WriteLog("Preferences not saved: " & Error$)
End Sub

' Cache folder ($(user)/calcmap/), created if needed. "" if impossible.
' On first launch, the cache of the former ColorierCarte extension
' (downloaded maps, palettes, preferences) is copied.
Function MapsFolder() As String
	Dim oSFA As Object, sUrl As String, sOld As String
	On Error GoTo Failed
	sUrl = CreateUnoService("com.sun.star.util.PathSubstitution") _
		.substituteVariables("$(user)/" & MAPS_FOLDER, True)
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If Not oSFA.exists(sUrl) Then
		sOld = CreateUnoService("com.sun.star.util.PathSubstitution") _
			.substituteVariables("$(user)/" & OLD_FOLDER, True)
		If oSFA.exists(sOld) Then
			oSFA.copy(sOld, sUrl)
			WriteLog("ColorierCarte cache taken over: " & ConvertFromURL(sOld))
		Else
			oSFA.createFolder(sUrl)
		End If
	End If
	MapsFolder = sUrl & "/"
	Exit Function
Failed:
	WriteLog("ERROR: maps folder not accessible (" & Error$ & ")")
	MapsFolder = ""
End Function

' List of maps: online catalog (copied into the cache) or, without a network,
' maps present in the cache.
Function LoadCatalog(sFolder As String, bOnline As Boolean, Optional bCacheOnly As Variant) As Variant
	Dim sCat As String, sHash As String, aMaps As Variant, oSFA As Object, sAggregate As String
	Dim bDownload As Boolean
	On Error GoTo Failed
	bOnline = False
	aMaps = Array()
	bDownload = IsMissing(bCacheOnly)
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	sCat = sFolder & CATALOG_NAME
	If bDownload Then
		WriteLog("Catalog: " & CatalogUrl())
		If DownloadFile(CatalogUrl(), sCat & ".part", sHash, Nothing, 0) Then
			If IsArray(JsonRead(ReadText(sCat & ".part"))) Then
				ReplaceFile(sCat & ".part", sCat)
			Else
				WriteLog("Unreadable catalog: using the cached copy")
				DeleteFile(sCat & ".part")
			End If
		End If
	End If
	' Cached copy, or else the catalog shipped with the extension
	sAggregate = sCat
	If Not oSFA.exists(sAggregate) Then
		sAggregate = ExtensionFolder() & "catalog.json"
		WriteLog("No cached catalog: using the extension's catalog")
	End If
	aMaps = ReadSources(sAggregate, sFolder, bDownload, bOnline)
	If bOnline Then
		WriteLog("Online catalog: " & (UBound(aMaps) + 1) & " map(s)")
	Else
		aMaps = CachedMaps(aMaps, sFolder)
		WriteLog("Offline: " & (UBound(aMaps) + 1) & " map(s) in cache")
	End If
	' After CachedMaps: the cached file of a discarded map is not
	' offered as an unknown local map
	LoadCatalog = FilterMaps(aMaps, MaxZones())
	Exit Function
Failed:
	WriteLog("ERROR loading the catalog: " & Error$)
	LoadCatalog = Array()
End Function

' Maps of an aggregated catalog ({"sources": [{id, title, url}…]}): each
' source is a catalog of maps, downloaded if bDownload (cached copy
' source-<id>.json otherwise). A catalog of maps ({"maps": …}) is
' also accepted directly. bOnline: at least one source up to date online.
Function ReadSources(sUrl As String, sFolder As String, bDownload As Boolean, bOnline As Boolean) As Variant
	Dim oRoot As Variant, aSources As Variant, i As Long, k As Long, sId As String, sSourceUrl As String
	Dim sLocal As String, sHash As String, aMaps As Variant, aRes() As Variant, n As Long
	Dim oIds As Object
	On Error GoTo Failed
	ReadSources = Array()
	oRoot = JsonRead(ReadText(sUrl))
	If IsArray(JsonField(oRoot, "maps")) Then
		ReadSources = ReadCatalog(sUrl, sFolder, "", bDownload)
		bOnline = bDownload And (UBound(ReadSources) >= 0)
		Exit Function
	End If
	oIds = com.sun.star.container.EnumerableMap.create("string", "boolean")
	aSources = JsonElements(JsonField(oRoot, "sources"))
	For i = 0 To UBound(aSources)
		sId = JsonText(aSources(i), "id")
		sSourceUrl = JsonText(aSources(i), "url")
		If Not SafeId(sId) Or Not IsAllowedUrl(sSourceUrl) Then
			WriteLog("Catalog: source " & (i + 1) & " skipped (invalid id or URL)")
		Else
			sLocal = sFolder & "source-" & sId & ".json"
			aMaps = Array()
			If bDownload Then
				If DownloadFile(sSourceUrl, sLocal & ".part", sHash, Nothing, 0) Then
					aMaps = ReadCatalog(sLocal & ".part", sFolder, sId, True)
					If UBound(aMaps) >= 0 Then
						ReplaceFile(sLocal & ".part", sLocal)
						bOnline = True
					Else
						DeleteFile(sLocal & ".part")
					End If
				End If
			End If
			If UBound(aMaps) < 0 Then aMaps = ReadCatalog(sLocal, sFolder, sId, False)
			WriteLog("Source " & sId & ": " & (UBound(aMaps) + 1) & " map(s)")
			For k = 0 To UBound(aMaps)
				If oIds.containsKey(aMaps(k)(C_ID)) Then
					WriteLog("  map '" & aMaps(k)(C_ID) & "' already provided by another source: skipped")
				Else
					If aMaps(k)(C_ID) <> "" Then oIds.put(aMaps(k)(C_ID), True)
					ReDim Preserve aRes(n)
					aRes(n) = aMaps(k)
					n = n + 1
				End If
			Next k
		End If
	Next i
	If n > 0 Then ReadSources = aRes
	Exit Function
Failed:
	WriteLog("ERROR reading the catalog sources: " & Error$)
	If n > 0 Then ReadSources = aRes
End Function

' Removes maps with more than nMax zones (unknown count: map kept)
Function FilterMaps(aMaps As Variant, nMax As Long) As Variant
	Dim aRes() As Variant, n As Long, i As Long
	FilterMaps = Array()
	For i = 0 To UBound(aMaps)
		If aMaps(i)(C_OBJECTS) > nMax Then
			WriteLog("Map not offered: " & aMaps(i)(C_TITLE) & " (" & aMaps(i)(C_OBJECTS) _
				& " zones, at most " & nMax & ")")
		Else
			ReDim Preserve aRes(n)
			aRes(n) = aMaps(i)
			n = n + 1
		End If
	Next i
	If n > 0 Then FilterMaps = aRes
End Function

' Extension settings (LibreOffice configuration, CalcMap.xcs)
Function CatalogUrl() As String
	CatalogUrl = Trim(ReadConfiguration(SETTINGS_NODE, "Catalog"))
	If CatalogUrl = "" Then CatalogUrl = DEFAULT_CATALOG
End Function

Function MaxZones() As Long
	Dim s As String
	s = ReadConfiguration(SETTINGS_NODE, "MaxZones")
	If Val(s) > 0 Then MaxZones = CLng(Val(s)) Else MaxZones = DEFAULT_MAX_ZONES
End Function

' URL accepted for a catalog, a map or a preview: https, or file
' (company network share)
Function IsAllowedUrl(s As String) As Boolean
	IsAllowedUrl = (LCase(Left(s, 8)) = "https://") Or (LCase(Left(s, 8)) = "file:///")
End Function

' Source id usable in a file name
Function SafeId(s As String) As Boolean
	Dim k As Integer
	SafeId = False
	If s = "" Or Len(s) > 64 Then Exit Function
	For k = 1 To Len(s)
		If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-", Mid(s, k, 1), 0) = 0 Then Exit Function
	Next k
	SafeId = True
End Function

' Reads a local maps.json; skips incomplete or suspicious entries.
' Texts in the interface language: see CatalogTexts
Function ReadCatalog(sUrl As String, sFolder As String, sSource As String, bDownload As Boolean) As Variant
	Dim oSFA As Object, oRoot As Variant, aList As Variant, oMapEntry As Variant, oTexts As Variant
	Dim aRes() As Variant, n As Long, i As Long, sFile As String, sMapUrl As String, sPreview As String
	On Error GoTo Failed
	ReadCatalog = Array()
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If Not oSFA.exists(sUrl) Then Exit Function
	oRoot = JsonRead(ReadText(sUrl))
	aList = JsonElements(JsonField(oRoot, "maps"))
	oTexts = CatalogTexts(oRoot, sFolder, sSource, bDownload)
	For i = 0 To UBound(aList)
		oMapEntry = aList(i)
		sFile = JsonText(oMapEntry, "filename")
		sMapUrl = JsonText(oMapEntry, "url")
		sPreview = JsonText(oMapEntry, "preview")
		If Not IsAllowedUrl(sPreview) Then sPreview = ""
		If SafeFileName(sFile) And IsAllowedUrl(sMapUrl) Then
			ReDim Preserve aRes(n)
			aRes(n) = NewMap(JsonText(oMapEntry, "id"), MapText(oTexts, oMapEntry, "title"), _
				MapText(oTexts, oMapEntry, "description"), sMapUrl, sFile, JsonNumber(oMapEntry, "size"), _
				LCase(JsonText(oMapEntry, "sha256")), oSFA.exists(sFolder & sFile), _
				JsonNumber(oMapEntry, "objects"), sPreview, LCase(JsonText(oMapEntry, "preview_sha256")), sSource, _
				CodeExamples(JsonField(oMapEntry, "code_examples")))
			n = n + 1
		Else
			WriteLog("Catalog: entry " & (i + 1) & " skipped (invalid file or URL)")
		End If
	Next i
	If n > 0 Then ReadCatalog = aRes
	Exit Function
Failed:
	WriteLog("ERROR reading the catalog: " & Error$)
	ReadCatalog = Array()
End Function

' Texts of the maps in the interface language ({id: {title, description}}),
' from the file announced by the catalog ("translations": URL with {lang},
' "languages": list of the available ones). Downloaded if bDownload, cached
' copy otherwise; Empty if there is none: the catalog's own texts are used
Function CatalogTexts(oRoot As Variant, sFolder As String, sSource As String, bDownload As Boolean) As Variant
	Dim oSFA As Object, aLangs As Variant, sLang As String, sUrl As String, sLocal As String, sHash As String
	Dim i As Long, bFound As Boolean
	On Error GoTo Failed
	CatalogTexts = Empty
	sLang = UiLanguage()
	sUrl = JsonText(oRoot, "translations")
	aLangs = JsonElements(JsonField(oRoot, "languages"))
	For i = 0 To UBound(aLangs)
		If aLangs(i) = sLang Then bFound = True
	Next i
	If Not bFound Or Not SafeId(sLang) Then Exit Function
	sUrl = Replace(sUrl, "{lang}", sLang)
	If Not IsAllowedUrl(sUrl) Then Exit Function
	sLocal = sFolder & "texts-" & IIf(sSource = "", "catalog", sSource) & "-" & sLang & ".json"
	If bDownload Then
		If DownloadFile(sUrl, sLocal & ".part", sHash, Nothing, 0) Then
			If IsArray(JsonRead(ReadText(sLocal & ".part"))) Then
				ReplaceFile(sLocal & ".part", sLocal)
			Else
				WriteLog("Unreadable map texts: " & sUrl)
				DeleteFile(sLocal & ".part")
			End If
		End If
	End If
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sLocal) Then CatalogTexts = JsonRead(ReadText(sLocal))
	Exit Function
Failed:
	WriteLog("ERROR reading the map texts: " & Error$)
	CatalogTexts = Empty
End Function

' Text of a map: file of the language, else the catalog (title_<lang>,
' then the default text in English)
Function MapText(oTexts As Variant, oMapEntry As Variant, sKey As String) As String
	Dim s As String
	s = JsonText(JsonField(oTexts, JsonText(oMapEntry, "id")), sKey)
	If s = "" Then s = JsonTranslatedText(oMapEntry, sKey)
	MapText = s
End Function

' Catalog maps present in the cache, plus the other .odg files of the cache

Function CachedMaps(aCatalog As Variant, sFolder As String) As Variant
	Dim oSFA As Object, aRes() As Variant, n As Long, i As Long, k As Long
	Dim aFiles As Variant, aParts As Variant, sName As String, bKnown As Boolean
	On Error GoTo Failed
	CachedMaps = Array()
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	For i = 0 To UBound(aCatalog)
		If aCatalog(i)(C_LOCALE) Then
			ReDim Preserve aRes(n)
			aRes(n) = aCatalog(i)
			n = n + 1
		End If
	Next i
	aFiles = oSFA.getFolderContents(sFolder, False)
	For i = 0 To UBound(aFiles)
		aParts = Split(ConvertFromURL(aFiles(i)), GetPathSeparator())
		sName = aParts(UBound(aParts))
		If LCase(Right(sName, 4)) = ".odg" Then
			bKnown = False
			For k = 0 To n - 1
				If aRes(k)(C_FILE) = sName Then bKnown = True
			Next k
			If Not bKnown Then
				ReDim Preserve aRes(n)
				aRes(n) = NewMap("", sName, Tr("cache.local_map"), "", _
					sName, oSFA.getSize(aFiles(i)), "", True)
				n = n + 1
			End If
		End If
	Next i
	If n > 0 Then CachedMaps = aRes
	Exit Function
Failed:
	WriteLog("ERROR reading the cache: " & Error$)
	If n > 0 Then CachedMaps = aRes
End Function

Function NewMap(sId, sTitle, sDescription, sUrl, sFile, nSize, sSha256, bLocale, _
		Optional nObjects, Optional sPreview, Optional sPreviewSha, Optional sSource, Optional sCodes) As Variant
	Dim v(4) As Variant
	v(0) = 0 : v(1) = "" : v(2) = "" : v(3) = "" : v(4) = ""
	If Not IsMissing(nObjects) Then v(0) = nObjects
	If Not IsMissing(sPreview) Then v(1) = sPreview
	If Not IsMissing(sPreviewSha) Then v(2) = sPreviewSha
	If Not IsMissing(sSource) Then v(3) = sSource
	If Not IsMissing(sCodes) Then v(4) = sCodes
	NewMap = Array(sId, sTitle, sDescription, sUrl, sFile, nSize, sSha256, bLocale, _
		v(0), v(1), v(2), v(3), v(4))
End Function

' Code examples from the catalog (JSON array of texts): "01, 02, 03"
Function CodeExamples(oList As Variant) As String
	Dim aCodes As Variant, i As Integer, s As String
	aCodes = JsonElements(oList)
	For i = 0 To UBound(aCodes)
		If VarType(aCodes(i)) = 8 And i < 6 Then s = s & IIf(s = "", "", ", ") & Left(aCodes(i), 12)
	Next i
	CodeExamples = s
End Function

' Local file of a map preview (SVG, PNG or JPEG), downloaded into
' the cache if needed and checked by its hash; "" if unavailable
Function LocalPreview(aMap As Variant) As String
	Dim oSFA As Object, sUrl As String, sExt As String, sFolder As String, sLocal As String, sHash As String
	On Error GoTo Failed
	LocalPreview = ""
	sUrl = aMap(C_PREVIEW)
	If sUrl = "" Or Not SafeId(aMap(C_ID)) Then Exit Function
	sExt = LCase(Right(sUrl, 4))
	If sExt <> ".svg" And sExt <> ".png" And sExt <> ".jpg" Then Exit Function
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	sFolder = MapsFolder() & PREVIEWS_FOLDER & "/"
	If Not oSFA.exists(sFolder) Then oSFA.createFolder(sFolder)
	' Name tied to the hash: a new preview is downloaded again
	sLocal = sFolder & aMap(C_ID) & IIf(aMap(C_PREVIEW_SHA) = "", "", "-" & Left(aMap(C_PREVIEW_SHA), 12)) & sExt
	If oSFA.exists(sLocal) Then
		LocalPreview = sLocal
		Exit Function
	End If
	If DownloadFile(sUrl, sLocal & ".part", sHash, Nothing, 0) Then
		If aMap(C_PREVIEW_SHA) = "" Or sHash = "" Or sHash = aMap(C_PREVIEW_SHA) Then
			ReplaceFile(sLocal & ".part", sLocal)
			LocalPreview = sLocal
		Else
			WriteLog("Preview of " & aMap(C_ID) & ": wrong hash, skipped")
			DeleteFile(sLocal & ".part")
		End If
	End If
	Exit Function
Failed:
	WriteLog("Preview of " & aMap(C_ID) & " unavailable: " & Error$)
	LocalPreview = ""
End Function

' Readable image (graphic) of a file; Nothing if impossible
Function GraphicOf(sUrl As String) As Object
	Dim aSource(0) As New com.sun.star.beans.PropertyValue
	On Error GoTo Failed
	GraphicOf = Nothing
	If sUrl = "" Then Exit Function
	aSource(0).Name = "URL"
	aSource(0).Value = sUrl
	GraphicOf = CreateUnoService("com.sun.star.graphic.GraphicProvider").queryGraphic(aSource())
	Exit Function
Failed:
	GraphicOf = Nothing
End Function

' Simple .odg file name (no path: the catalog comes from the network)
Function SafeFileName(s As String) As Boolean
	Dim k As Integer, c As String
	SafeFileName = False
	If Len(s) < 5 Or LCase(Right(s, 4)) <> ".odg" Or Left(s, 1) = "." Then Exit Function
	For k = 1 To Len(s)
		c = Mid(s, k, 1)
		If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._-", c, 0) = 0 Then Exit Function
	Next k
	SafeFileName = True
End Function

' Returns the URL of the cached map, downloaded or updated if its
' hash differs from the catalog's. "" if no usable version.
Function LocalMap(aMap As Variant, sFolder As String, bOnline As Boolean, oDocCalc As Object) As String
	Dim oSFA As Object, sLocal As String, sHash As String, bExists As Boolean
	Dim oIndicator As Object
	On Error GoTo Failed
	LocalMap = ""
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	sLocal = sFolder & aMap(C_FILE)
	bExists = oSFA.exists(sLocal)

	If Not bOnline Or aMap(C_URL) = "" Then
		If bExists Then
			LocalMap = sLocal
		Else
			Warn(Tr("msg.not_offline", aMap(C_TITLE)))
		End If
		Exit Function
	End If

	If bExists Then
		If IsUpToDate(sLocal, aMap) Then
			WriteLog("Cached map up to date: no download")
			LocalMap = sLocal
			Exit Function
		End If
		WriteLog("Cached map differs from the catalog: updating")
	Else
		WriteLog("Map not in cache: downloading")
	End If

	oIndicator = StatusIndicator(oDocCalc)
	If DownloadFile(aMap(C_URL), sLocal & ".part", sHash, oIndicator, aMap(C_SIZE)) Then
		If aMap(C_SHA256) = "" Or sHash = "" Or sHash = aMap(C_SHA256) Then
			ReplaceFile(sLocal & ".part", sLocal)
			WriteLog("Map downloaded (SHA-256 " & sHash & ")")
			LocalMap = sLocal
			Exit Function
		End If
		WriteLog("Wrong download hash: " & sHash _
			& " instead of " & aMap(C_SHA256))
		DeleteFile(sLocal & ".part")
	End If

	If bExists Then
		WriteLog("Download impossible: using the old cached version")
		LocalMap = sLocal
	Else
		Warn(Tr("msg.download_failed", aMap(C_TITLE)))
	End If
	Exit Function
Failed:
	WriteLog("ERROR on the map " & aMap(C_FILE) & ": " & Error$)
	If bExists Then LocalMap = sLocal
End Function

' Compares the cached file with the catalog: SHA-256, or size if the hash
' computation is unavailable
Function IsUpToDate(sLocal As String, aMap As Variant) As Boolean
	Dim sHash As String, bSame As Boolean
	On Error GoTo Failed
	IsUpToDate = False
	If aMap(C_SHA256) <> "" Then
		sHash = FileSha256(sLocal)
		If sHash <> "" Then
			WriteLog("SHA-256 local " & sHash & " ; catalog " & aMap(C_SHA256))
			IsUpToDate = (sHash = aMap(C_SHA256))
			Exit Function
		End If
	End If
	If aMap(C_SIZE) > 0 Then
		bSame = (CreateUnoService("com.sun.star.ucb.SimpleFileAccess").getSize(sLocal) = aMap(C_SIZE))
		WriteLog("Comparison by size: " & IIf(bSame, "same", "different"))
		IsUpToDate = bSame
	Else
		IsUpToDate = True
	End If
	Exit Function
Failed:
	IsUpToDate = False
End Function

' ---------------------------------------------------------------------------
' Download and hashes
' ---------------------------------------------------------------------------

' Copies sSource (https or file) to sDest while computing its SHA-256.
' False on failure (no network, etc.); sDest is then deleted.
Function DownloadFile(sSource As String, sDest As String, sHash As String, _
		oIndicator As Object, ByVal nSize As Double) As Boolean
	Dim oSFA As Object, oEntry As Object, oOut As Object, oCtx As Object
	Dim aBytes(), n As Long, nRead As Double
	On Error GoTo Failed
	DownloadFile = False
	sHash = ""
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sDest) Then oSFA.kill(sDest)
	oEntry = oSFA.openFileRead(sSource)
	oOut = oSFA.openFileWrite(sDest)
	oCtx = Sha256Context()
	If Not IsNull(oIndicator) Then oIndicator.start(Tr("progress.download"), 100)
	Do
		n = oEntry.readBytes(aBytes, 65536)
		If n > 0 Then
			oOut.writeBytes(aBytes)
			If Not IsNull(oCtx) Then oCtx.updateDigest(aBytes)
			nRead = nRead + n
			If Not IsNull(oIndicator) And nSize > 0 Then
				oIndicator.setValue(IIf(nRead < nSize, Int(nRead / nSize * 100), 100))
			End If
		End If
	Loop While n > 0
	oEntry.closeInput()
	oOut.closeOutput()
	If Not IsNull(oCtx) Then sHash = ToHex(oCtx.finalizeDigestAndDispose())
	If Not IsNull(oIndicator) Then oIndicator.end()
	DownloadFile = True
	Exit Function
Failed:
	WriteLog("Download impossible: " & sSource & ": " & Error$)
	On Error Resume Next
	If Not IsNull(oIndicator) Then oIndicator.end()
	If Not IsNull(oEntry) Then oEntry.closeInput()
	If Not IsNull(oOut) Then oOut.closeOutput()
	If oSFA.exists(sDest) Then oSFA.kill(sDest)
	DownloadFile = False
End Function

' SHA-256 of a file in lowercase hexadecimal ("" if unavailable)
Function FileSha256(sUrl As String) As String
	Dim oEntry As Object, oCtx As Object, aBytes(), n As Long
	On Error GoTo Failed
	FileSha256 = ""
	oCtx = Sha256Context()
	If IsNull(oCtx) Then Exit Function
	oEntry = CreateUnoService("com.sun.star.ucb.SimpleFileAccess").openFileRead(sUrl)
	Do
		n = oEntry.readBytes(aBytes, 65536)
		If n > 0 Then oCtx.updateDigest(aBytes)
	Loop While n > 0
	oEntry.closeInput()
	FileSha256 = ToHex(oCtx.finalizeDigestAndDispose())
	Exit Function
Failed:
	WriteLog("SHA-256 impossible for " & ConvertFromURL(sUrl) & ": " & Error$)
	On Error Resume Next
	If Not IsNull(oEntry) Then oEntry.closeInput()
	FileSha256 = ""
End Function

' SHA-256 context provided by LibreOffice (NSS); Nothing if unavailable
Function Sha256Context() As Object
	On Error GoTo Failed
	Sha256Context = CreateUnoService("com.sun.star.xml.crypto.NSSInitializer") _
		.getDigestContext(com.sun.star.xml.crypto.DigestID.SHA256, Array())
	Exit Function
Failed:
	WriteLog("SHA-256 computation unavailable: " & Error$)
	Sha256Context = Nothing
End Function

Function ToHex(aBytes As Variant) As String
	Dim i As Long, s As String
	For i = LBound(aBytes) To UBound(aBytes)
		s = s & Right("0" & LCase(Hex(aBytes(i) And 255)), 2)
	Next i
	ToHex = s
End Function

Function ReadText(sUrl As String) As String
	Dim oText As Object, s As String
	On Error GoTo Failed
	oText = CreateUnoService("com.sun.star.io.TextInputStream")
	oText.setInputStream(CreateUnoService("com.sun.star.ucb.SimpleFileAccess").openFileRead(sUrl))
	oText.setEncoding("UTF-8")
	Do While Not oText.isEOF()
		s = s & oText.readLine() & Chr(10)
	Loop
	oText.closeInput()
	ReadText = s
	Exit Function
Failed:
	WriteLog("ERROR reading " & ConvertFromURL(sUrl) & ": " & Error$)
	ReadText = ""
End Function

Sub ReplaceFile(sFrom As String, sTo As String)
	Dim oSFA As Object
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sTo) Then oSFA.kill(sTo)
	oSFA.move(sFrom, sTo)
End Sub

Sub DeleteFile(sUrl As String)
	Dim oSFA As Object
	On Error Resume Next
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sUrl) Then oSFA.kill(sUrl)
End Sub

' Calc window, parent of the dialog (Nothing if unavailable)
Function ParentWindow(oDocCalc As Object) As Object
	On Error GoTo Failed
	ParentWindow = oDocCalc.getCurrentController().getFrame().getContainerWindow()
	Exit Function
Failed:
	ParentWindow = Nothing
End Function

' Progress bar of the Calc window (Nothing if unavailable)
Function StatusIndicator(oDocCalc As Object) As Object
	On Error GoTo Failed
	StatusIndicator = oDocCalc.getCurrentController().getFrame().createStatusIndicator()
	Exit Function
Failed:
	StatusIndicator = Nothing
End Function

Function ReadableSize(nSize As Double) As String
	If nSize >= 1048576 Then
		ReadableSize = Tr("size.mb", Format(nSize / 1048576, "0.0"))
	ElseIf nSize > 0 Then
		ReadableSize = Tr("size.kb", Format(nSize / 1024, "0"))
	Else
		ReadableSize = ""
	End If
End Function

' ---------------------------------------------------------------------------
' Dialog to choose the map, the palette, the columns and the options.
' Returns 1 (OK), 0 (cancel) or 3 (other file); iChoice and iPalette
' receive the chosen indexes, aOptions the options.
' bConfig: configuration of an already inserted map (base map, insertion
' and "Other file…" cannot be changed).
' ---------------------------------------------------------------------------
Function MapDialog(aMaps As Variant, aPalettes As Variant, sState As String, _
		sPrefMap As String, sPrefPalette As String, iChoice As Integer, iPalette As Integer, _
		bInsert As Boolean, aOptions As Variant, aColumns As Variant, oParent As Object, _
		bConfig As Boolean) As Integer
	Dim oModel As Object, oList As Object, oPalList As Object
	Dim oListListener As Object, oPalListener As Object, oButtonListener As Object, oChecksListener As Object
	Dim aLabels() As String, aPalNames() As String, i As Long, iDefault As Integer, nLast As Long
	Dim iPalDefault As Integer, aFonts As Variant, sCheck As Variant, oClassesListener As Object, oCodeListener As Object
	On Error GoTo Failed
	MapDialog = 0
	iChoice = -1
	iPalette = -1
	aDialogMaps = aMaps
	aDialogPalettes = aPalettes

	nLast = UBound(aMaps)
	If nLast >= 0 Then ReDim aLabels(nLast) As String
	iDefault = -1
	For i = 0 To nLast
		aLabels(i) = aMaps(i)(C_TITLE)
		If Not aMaps(i)(C_LOCALE) Then
			aLabels(i) = aLabels(i) & "  - " & Tr("dlg.to_download", ReadableSize(aMaps(i)(C_SIZE)))
		End If
		If sPrefMap <> "" And (aMaps(i)(C_ID) = sPrefMap Or aMaps(i)(C_FILE) = sPrefMap) Then
			iDefault = i
		ElseIf iDefault < 0 And aMaps(i)(C_ID) = DEFAULT_MAP Then
			iDefault = i
		End If
	Next i
	If iDefault < 0 Then iDefault = 0

	ReDim aPalNames(UBound(aPalettes)) As String
	For i = 0 To UBound(aPalettes)
		aPalNames(i) = aPalettes(i)(P_NAME)
	Next i
	iPalDefault = PaletteIndex(aPalettes, sPrefPalette)
	aFonts = AvailableFonts(Array(aOptions(O_FONT_VALUES), aOptions(O_FONT_LABELS)))

	oModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
	oModel.Title = IIf(bConfig, Tr("dlg.config_title"), Tr("app.title"))
	oModel.Height = 340
	oModel.Width = 460
	' Top: Map (base map, description, insertion, preview); below, side by
	' side like the columns to choose: Code, Values, Labels (narrow
	' frames: one list or one row of buttons per line); then Legend
	AddFrame(oModel, "grpMap", 4, 4, 452, 92, Tr("dlg.grp_map"))
	AddControl(oModel, "ListBox", "lstMap", 12, 16, 346, 14, _
		Array("Dropdown", "LineCount", "StringItemList", "Enabled"), Array(True, 12, aLabels, Not bConfig))
	If bConfig Then
		' Configuration: the base map does not change, its name is enough
		AddControl(oModel, "FixedText", "lblMapTitle", 12, 18, 346, 12, _
			Array("Label", "FontWeight"), Array(IIf(nLast >= 0, aLabels(0), ""), com.sun.star.awt.FontWeight.BOLD))
	End If
	AddControl(oModel, "FixedText", "lblDescription", 12, 33, 346, 36, Array("MultiLine"), Array(True))
	AddControl(oModel, "CheckBox", "chkInsert", 12, 72, 346, 20, _
		Array("Label", "State", "Enabled", "MultiLine"), Array(Tr("dlg.insert_in_calc"), IIf(bInsert, 1, 0), Not bConfig, True))
	' Preview of the chosen map (default image if none)
	AddControl(oModel, "ImageControl", "imgPreview", 364, 14, 84, 70, _
		Array("ScaleImage", "ScaleMode", "Border"), Array(True, com.sun.star.awt.ImageScaleMode.ISOTROPIC, 1))
	' Code: its column, and the map codes (examples) to compare,
	' with an icon: found in the column, or not
	AddFrame(oModel, "grpCode", 4, 100, 148, 182, Tr("dlg.grp_code"))
	AddColumnList(oModel, "Code", 12, 112, aColumns, aOptions(O_COL_CODE), Tr("dlg.column"), 132)
	AddControl(oModel, "ImageControl", "imgCodes", 12, 142, 10, 10, _
		Array("ScaleImage", "ScaleMode", "Border"), Array(True, com.sun.star.awt.ImageScaleMode.ISOTROPIC, 0))
	AddControl(oModel, "FixedText", "lblMapCodes", 26, 142, 118, 20, _
		Array("Label", "TextColor", "MultiLine", "HelpText"), Array("", RGB(96, 96, 96), True, Tr("dlg.map_codes_help")))
	' Values: column, colors (palette, preview, classes), display,
	' separated by lines
	AddFrame(oModel, "grpValues", 156, 100, 148, 182, Tr("dlg.grp_values"))
	AddColumnList(oModel, "Value", 164, 112, aColumns, aOptions(O_COL_VALUE), Tr("dlg.column"), 132)
	AddControl(oModel, "FixedLine", "linValues1", 164, 140, 132, 4, Array(), Array())
	AddControl(oModel, "FixedText", "lblPalette", 164, 149, 76, 10, Array("Label"), Array(Tr("dlg.colors")))
	AddControl(oModel, "Button", "btnClasses", 246, 146, 50, 14, _
		Array("Label", "PushButtonType", "HelpText"), Array(Tr("dlg.classes"), 0, Tr("dlg.classes_help")))
	AddControl(oModel, "ListBox", "lstPalette", 164, 162, 132, 14, _
		Array("Dropdown", "LineCount", "StringItemList"), Array(True, 16, aPalNames))
	For i = 0 To PREVIEW_SHADES - 1
		AddControl(oModel, "FixedText", "shade" & i, PREVIEW_X, 179, 10, 12, Array("Border"), Array(2))
	Next i
	AddControl(oModel, "FixedText", "lblClasses", 164, 193, 132, 10, _
		Array("Label", "TextColor"), Array("", RGB(96, 96, 96)))
	sDialogClasses = aOptions(O_CLASSES)
	AddControl(oModel, "FixedLine", "linValues2", 164, 205, 132, 4, Array(), Array())
	AddTextRow(oModel, "Values", 164, 211, Tr("dlg.show_values"), aOptions(O_VALUES), _
		aOptions(O_SIZE_VALUES), aFonts, aOptions(O_FONT_VALUES), aOptions(O_STYLE_VALUES), _
		HasOutline(aOptions(O_OUTLINE_VALUES)))
	' Labels: column, display (at the level of the values one)
	AddFrame(oModel, "grpLabels", 308, 100, 148, 182, Tr("dlg.grp_labels"))
	AddColumnList(oModel, "Label", 316, 112, aColumns, aOptions(O_COL_LABEL), Tr("dlg.column"), 132)
	AddControl(oModel, "FixedLine", "linLabels", 316, 140, 132, 4, Array(), Array())
	AddTextRow(oModel, "Labels", 316, 211, Tr("dlg.show_labels"), aOptions(O_LABELS), _
		aOptions(O_SIZE_LABELS), aFonts, aOptions(O_FONT_LABELS), aOptions(O_STYLE_LABELS), _
		HasOutline(aOptions(O_OUTLINE_LABELS)))
	' Legend, at the very bottom
	AddFrame(oModel, "grpLegend", 4, 286, 452, 28, Tr("dlg.grp_legend"))
	AddControl(oModel, "CheckBox", "chkLegend", 12, 299, 140, 12, _
		Array("Label", "State"), Array(Tr("dlg.show_legend"), IIf(aOptions(O_LEGEND), 1, 0)))
	AddControl(oModel, "CheckBox", "chkLegendValues", 164, 299, 284, 12, _
		Array("Label", "State", "Enabled"), Array(Tr("dlg.legend_values"), IIf(aOptions(O_LEGEND_VALUES), 1, 0), _
		aOptions(O_LEGEND)))
	' Buttons, and the catalog state between "Other file…" and OK
	AddControl(oModel, "Button", "btnOther", 6, 320, 70, 14, _
		Array("Label", "PushButtonType"), Array(Tr("dlg.other_file"), 0))
	AddControl(oModel, "FixedText", "lblState", 80, 318, 266, 20, _
		Array("MultiLine", "Label", "TextColor"), Array(True, sState, RGB(96, 96, 96)))
	AddControl(oModel, "Button", "btnOk", 350, 320, 50, 14, _
		Array("Label", "PushButtonType", "DefaultButton", "Enabled"), Array(Tr("dlg.ok"), 1, True, nLast >= 0))
	AddControl(oModel, "Button", "btnCancel", 404, 320, 50, 14, _
		Array("Label", "PushButtonType"), Array(Tr("dlg.cancel"), 2))
	If nLast >= 0 Then oModel.getByName("lstMap").SelectedItems = Array(CInt(iDefault))
	oModel.getByName("lstPalette").SelectedItems = Array(CInt(iPalDefault))

	oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
	oDialog.setModel(oModel)
	oList = oDialog.getControl("lstMap")
	oListListener = CreateUnoListener("MapList_", "com.sun.star.awt.XItemListener")
	oList.addItemListener(oListListener)
	oCodeListener = CreateUnoListener("CodeList_", "com.sun.star.awt.XItemListener")
	oDialog.getControl("lstCode").addItemListener(oCodeListener)
	oPalList = oDialog.getControl("lstPalette")
	oPalListener = CreateUnoListener("PaletteList_", "com.sun.star.awt.XItemListener")
	oPalList.addItemListener(oPalListener)
	oButtonListener = CreateUnoListener("OtherFile_", "com.sun.star.awt.XActionListener")
	oDialog.getControl("btnOther").addActionListener(oButtonListener)
	oClassesListener = CreateUnoListener("ClassesButton_", "com.sun.star.awt.XActionListener")
	oDialog.getControl("btnClasses").addActionListener(oClassesListener)
	oChecksListener = CreateUnoListener("TextChecks_", "com.sun.star.awt.XItemListener")
	For Each sCheck In Array("chkValues", "chkLabels", "chkLegend")
		oDialog.getControl(sCheck).addItemListener(oChecksListener)
	Next sCheck
	' Without an explicitly created peer, execute() returns -1 without showing anything
	' (seen with LibreOffice 25.2 on Wayland)
	oDialog.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), oParent)
	If nLast >= 0 Then ShowDescription(iDefault)
	ShowPalette(iPalDefault)
	ShowClassesState()
	If bConfig Then
		oDialog.getControl("btnOther").setVisible(False)
		oDialog.getControl("lstMap").setVisible(False)
		oDialog.getControl("chkInsert").setVisible(False)
		oDialog.getControl("lblDescription").getModel().Label = Tr("dlg.config_hint")
	End If

	MapDialog = oDialog.execute()
	iChoice = oList.getSelectedItemPos()
	iPalette = oPalList.getSelectedItemPos()
	bInsert = (oDialog.getControl("chkInsert").getState() = 1)
	aOptions(O_COL_CODE) = ChosenColumn("Code", aOptions(O_COL_CODE))
	aOptions(O_COL_VALUE) = ChosenColumn("Value", aOptions(O_COL_VALUE))
	aOptions(O_COL_LABEL) = ChosenColumn("Label", aOptions(O_COL_LABEL))
	ReadTextRow("Values", aOptions, O_VALUES, O_SIZE_VALUES, O_FONT_VALUES, O_STYLE_VALUES, O_OUTLINE_VALUES)
	ReadTextRow("Labels", aOptions, O_LABELS, O_SIZE_LABELS, O_FONT_LABELS, O_STYLE_LABELS, _
		O_OUTLINE_LABELS)
	aOptions(O_LEGEND) = (oDialog.getControl("chkLegend").getState() = 1)
	aOptions(O_LEGEND_VALUES) = (oDialog.getControl("chkLegendValues").getState() = 1)
	aOptions(O_CLASSES) = sDialogClasses
	oDialog.getControl("btnClasses").removeActionListener(oClassesListener)
	For Each sCheck In Array("chkValues", "chkLabels", "chkLegend")
		oDialog.getControl(sCheck).removeItemListener(oChecksListener)
	Next sCheck
	If MapDialog = 0 Then WriteLog("Map choice dialog canceled")
	oList.removeItemListener(oListListener)
	oDialog.getControl("lstCode").removeItemListener(oCodeListener)
	oPalList.removeItemListener(oPalListener)
	oDialog.getControl("btnOther").removeActionListener(oButtonListener)
	If bKeepDialog Then Exit Function
	oDialog.dispose()
	oDialog = Nothing
	Exit Function
Failed:
	WriteLog("ERROR in the map choice dialog " & Err & " (line " & Erl & "): " & Error$)
	On Error Resume Next
	If Not IsNull(oDialog) Then oDialog.dispose()
	oDialog = Nothing
	MapDialog = 3          ' fallback: classic file picker
End Function

' Dialog text on four lines: check box; font and size (lists
' you can also type in); B I U buttons, like the formatting toolbar; outline
' check box, greyed out while the check box is not checked:
' chk<sSuffix>, cboFont<sSuffix>, cboSize<sSuffix>, btnBold<sSuffix>,
' chkOutline<sSuffix>…
Sub AddTextRow(oModel As Object, sSuffix As String, nX As Long, nY As Long, sLabel As String, _
		bChecked As Boolean, nSize As Integer, aFonts As Variant, sFont As String, sStyle As String, _
		bOutline As Boolean)
	Dim aSizes As Variant, i As Integer
	AddControl(oModel, "CheckBox", "chk" & sSuffix, nX, nY, 132, 12, _
		Array("Label", "State"), Array(sLabel, IIf(bChecked, 1, 0)))
	' Below: font and size (lists you can also type in), then B I U
	AddControl(oModel, "ComboBox", "cboFont" & sSuffix, nX, nY + 15, 82, 14, _
		Array("Dropdown", "LineCount", "StringItemList", "Text", "Autocomplete", "HelpText", "Enabled"), _
		Array(True, 16, aFonts, sFont, True, Tr("dlg.font"), bChecked))
	aSizes = Split(OFFERED_SIZES, "|")
	For i = 0 To UBound(aSizes)
		aSizes(i) = aSizes(i) & " pt"
	Next i
	AddControl(oModel, "ComboBox", "cboSize" & sSuffix, nX + 84, nY + 15, 48, 14, _
		Array("Dropdown", "LineCount", "StringItemList", "Text", "HelpText", "Enabled"), _
		Array(True, 12, aSizes, ValidSize(nSize) & " pt", Tr("dlg.text_size"), bChecked))
	' Bold, italic, underline: toggle buttons, letter in the style
	AddStyleButton(oModel, "btnBold" & sSuffix, nX, nY + 31, Tr("dlg.bold_short"), Tr("dlg.bold"), _
		InStr(sStyle, "B") > 0, bChecked, "FontWeight", com.sun.star.awt.FontWeight.BOLD)
	AddStyleButton(oModel, "btnItalic" & sSuffix, nX + 16, nY + 31, Tr("dlg.italic_short"), Tr("dlg.italic"), _
		InStr(sStyle, "I") > 0, bChecked, "FontSlant", com.sun.star.awt.FontSlant.ITALIC)
	AddStyleButton(oModel, "btnUnderline" & sSuffix, nX + 32, nY + 31, Tr("dlg.underline_short"), Tr("dlg.underline"), _
		InStr(sStyle, "U") > 0, bChecked, "FontUnderline", com.sun.star.awt.FontUnderline.SINGLE)
	' White or black outline, further below
	AddControl(oModel, "CheckBox", "chkOutline" & sSuffix, nX + 12, nY + 47, 120, 20, _
		Array("Label", "State", "Enabled", "MultiLine"), Array(Tr("dlg.outline"), IIf(bOutline, 1, 0), bChecked, True))
End Sub

Sub AddStyleButton(oModel As Object, sName As String, nX As Long, nY As Long, sLetter As String, _
		sHelp As String, bPressed As Boolean, bActive As Boolean, sProperty As String, vStyle As Variant)
	AddControl(oModel, "Button", sName, nX, nY, 14, 14, _
		Array("Label", "HelpText", "Toggle", "State", "Enabled", "PushButtonType", sProperty), _
		Array(sLetter, sHelp, True, IIf(bPressed, 1, 0), bActive, 0, vStyle))
End Sub

' Drop-down list of a column in its frame (lbl<sSuffix>, lst<sSuffix>)
Sub AddColumnList(oModel As Object, sSuffix As String, nX As Long, nY As Long, aColumns As Variant, _
		iDefault As Integer, sLabel As String, nWidth As Long)
	Dim i As Integer
	i = iDefault
	If i > UBound(aColumns) Then i = UBound(aColumns)
	If i < 0 Then i = 0
	' Label above the list (narrow frames)
	AddControl(oModel, "FixedText", "lbl" & sSuffix, nX, nY, nWidth, 10, Array("Label"), Array(sLabel))
	AddControl(oModel, "ListBox", "lst" & sSuffix, nX, nY + 11, nWidth, 14, _
		Array("Dropdown", "LineCount", "StringItemList"), Array(True, 12, aColumns))
	oModel.getByName("lst" & sSuffix).SelectedItems = Array(i)
End Sub

' Frame (title and border) grouping controls
Sub AddFrame(oModel As Object, sName As String, nX As Long, nY As Long, nWidth As Long, nHeight As Long, _
		sTitle As String)
	AddControl(oModel, "GroupBox", sName, nX, nY, nWidth, nHeight, Array("Label"), Array(sTitle))
End Sub

' Values of a dialog text row (check box, size, font, style)
Sub ReadTextRow(sSuffix As String, aOptions As Variant, kCheck As Integer, kSize As Integer, _
		kFont As Integer, kStyle As Integer, kOutline As Integer)
	Dim sStyle As String, sFont As String
	aOptions(kCheck) = (oDialog.getControl("chk" & sSuffix).getState() = 1)
	aOptions(kSize) = ValidSize(Val(oDialog.getControl("cboSize" & sSuffix).getText()))
	sFont = Trim(oDialog.getControl("cboFont" & sSuffix).getText())
	If sFont <> "" Then aOptions(kFont) = sFont
	If oDialog.getControl("btnBold" & sSuffix).getModel().State = 1 Then sStyle = sStyle & "B"
	If oDialog.getControl("btnItalic" & sSuffix).getModel().State = 1 Then sStyle = sStyle & "I"
	If oDialog.getControl("btnUnderline" & sSuffix).getModel().State = 1 Then sStyle = sStyle & "U"
	aOptions(kStyle) = sStyle
	aOptions(kOutline) = IIf(oDialog.getControl("chkOutline" & sSuffix).getState() = 1, _
		OUTLINE_ON, OUTLINE_NONE)
End Sub

Function ChosenColumn(sSuffix As String, iDefault As Integer) As Integer
	ChosenColumn = oDialog.getControl("lst" & sSuffix).getSelectedItemPos()
	If ChosenColumn < 0 Then ChosenColumn = iDefault
End Function

' Installed fonts (and those of aExtra), sorted
Function AvailableFonts(aExtra As Variant) As Variant
	Dim oNames As Object, aDesc As Variant, i As Long, aRes() As String, n As Long, oEnum As Object
	oNames = com.sun.star.container.EnumerableMap.create("string", "boolean")
	On Error GoTo Proceed
	aDesc = CreateUnoService("com.sun.star.awt.Toolkit").createScreenCompatibleDevice(16, 16).getFontDescriptors()
	For i = 0 To UBound(aDesc)
		If aDesc(i).Name <> "" Then oNames.put(aDesc(i).Name, True)
	Next i
Proceed:
	On Error Resume Next
	oNames.put(TEXT_FONT, True)
	For i = 0 To UBound(aExtra)
		If aExtra(i) <> "" Then oNames.put(aExtra(i), True)
	Next i
	oEnum = oNames.createKeyEnumeration(False)
	Do While oEnum.hasMoreElements()
		ReDim Preserve aRes(n) As String
		aRes(n) = oEnum.nextElement()
		n = n + 1
	Loop
	SortTexts(aRes)
	AvailableFonts = aRes
End Function

' Shell sort of texts, case insensitive
Sub SortTexts(a As Variant)
	Dim nGap As Long, i As Long, j As Long, s As String
	nGap = (UBound(a) + 1) \ 2
	Do While nGap > 0
		For i = nGap To UBound(a)
			s = a(i)
			j = i
			Do While j >= nGap
				If LCase(a(j - nGap)) <= LCase(s) Then Exit Do
				a(j) = a(j - nGap)
				j = j - nGap
			Loop
			a(j) = s
		Next i
		nGap = nGap \ 2
	Loop
End Sub

Sub AddControl(oModel As Object, sType As String, sName As String, nX As Long, nY As Long, _
		nWidth As Long, nHeight As Long, aNames As Variant, aValues As Variant)
	Dim oCtrl As Object, k As Integer
	oCtrl = oModel.createInstance("com.sun.star.awt.UnoControl" & sType & "Model")
	oCtrl.Name = sName          ' read by the listeners (not set by insertByName)
	oCtrl.PositionX = nX
	oCtrl.PositionY = nY
	oCtrl.Width = nWidth
	oCtrl.Height = nHeight
	For k = 0 To UBound(aNames)
		oCtrl.setPropertyValue(aNames(k), aValues(k))
	Next k
	oModel.insertByName(sName, oCtrl)
End Sub

Sub ShowDescription(ByVal i As Long)
	Dim aMap As Variant, s As String
	On Error Resume Next
	If i < 0 Or i > UBound(aDialogMaps) Then Exit Sub
	aMap = aDialogMaps(i)
	s = aMap(C_DESCRIPTION) & Chr(10) & aMap(C_FILE)
	If aMap(C_OBJECTS) > 0 Then s = s & " - " & Tr("dlg.zones", aMap(C_OBJECTS))
	If aMap(C_SIZE) > 0 Then s = s & " - " & ReadableSize(aMap(C_SIZE))
	s = s & " - " & IIf(aMap(C_LOCALE), Tr("dlg.downloaded"), Tr("dlg.will_download"))
	oDialog.getControl("lblDescription").getModel().Label = s
	oDialog.getControl("lblMapCodes").getModel().Label = _
		IIf(aMap(C_CODES) = "", "", Tr("dlg.map_codes", aMap(C_CODES) & ", ..."))
	CheckCodes(aMap)
	ShowPreview(aMap)
End Sub

' Preview of a map in the dialog; "no preview" image if the map
' has none or if it is unreadable
Sub ShowPreview(aMap As Variant)
	Dim oGraphic As Object
	On Error Resume Next
	oGraphic = GraphicOf(LocalPreview(aMap))
	If IsNull(oGraphic) Then oGraphic = GraphicOf(ExtensionFolder() & "icons/nopreview.svg")
	oDialog.getControl("imgPreview").getModel().Graphic = oGraphic
End Sub

' Palette preview: one box per class, or for a continuous gradient
' all the boxes, adjoining and without border
Sub ShowPalette(ByVal i As Long)
	Dim aColors As Variant, n As Integer, k As Integer, nWidth As Integer, oBox As Object
	Dim bContinuous As Boolean, nColor As Long
	On Error Resume Next
	If i < 0 Or i > UBound(aDialogPalettes) Then Exit Sub
	aColors = aDialogPalettes(i)(P_COLORS)
	bContinuous = aDialogPalettes(i)(P_CONTINUOUS)
	n = UBound(aColors) + 1
	If bContinuous Or n > PREVIEW_SHADES Then n = PREVIEW_SHADES
	nWidth = Int(PREVIEW_W / n)
	If nWidth > 20 Then nWidth = 20
	For k = 0 To PREVIEW_SHADES - 1
		oBox = oDialog.getControl("shade" & k)
		If k < n Then
			If bContinuous Then
				nColor = GradientColor(aColors, k / (n - 1))
			Else
				nColor = aColors(k)
			End If
			oBox.getModel().PositionX = PREVIEW_X + k * nWidth
			oBox.getModel().Width = nWidth
			oBox.getModel().Border = IIf(bContinuous, 0, 2)
			oBox.getModel().BackgroundColor = nColor
			oBox.setVisible(True)
		Else
			oBox.setVisible(False)
		End If
	Next k
End Sub

' Dialog listeners
' Code column changed: check the map codes
Sub CodeList_itemStateChanged(oEvent As Object)
	Dim i As Integer
	On Error Resume Next
	i = oDialog.getControl("lstMap").getSelectedItemPos()
	If i < 0 Then i = 0
	If i <= UBound(aDialogMaps) Then CheckCodes(aDialogMaps(i))
End Sub

Sub CodeList_disposing(oEvent As Object)
End Sub

' Icon before the map codes: green tick if all the example codes
' are in the chosen code column (lost leading zeros
' tolerated), otherwise a discreet amber triangle; nothing without examples
Sub CheckCodes(aMap As Variant)
	Dim oImage As Object, oColumn As Object, iCol As Integer, r As Long, sCode As String
	Dim aCodes As Variant, k As Integer, sMissing As String
	On Error Resume Next
	oImage = oDialog.getControl("imgCodes")
	If aMap(C_CODES) = "" Or Not IsArray(aDialogData) Then
		oImage.setVisible(False)
		Exit Sub
	End If
	iCol = oDialog.getControl("lstCode").getSelectedItemPos()
	oColumn = com.sun.star.container.EnumerableMap.create("string", "boolean")
	For r = 0 To UBound(aDialogData)
		If UBound(aDialogData(r)) >= iCol Then
			sCode = NormalizeCode(aDialogData(r)(iCol))
			If sCode <> "" Then oColumn.put(sCode, True)
		End If
	Next r
	aCodes = Split(aMap(C_CODES), ", ")
	For k = 0 To UBound(aCodes)
		If FoundKey(oColumn, aCodes(k)) = "" Then sMissing = sMissing & IIf(sMissing = "", "", ", ") & aCodes(k)
	Next k
	If sMissing = "" Then
		oImage.getModel().Graphic = GraphicOf(ExtensionFolder() & "icons/codes-ok.svg")
		oImage.getModel().HelpText = Tr("dlg.codes_ok")
	Else
		oImage.getModel().Graphic = GraphicOf(ExtensionFolder() & "icons/codes-attention.svg")
		oImage.getModel().HelpText = Tr("dlg.codes_missing", sMissing)
	End If
	oImage.setVisible(True)
End Sub

Sub MapList_itemStateChanged(oEvent As Object)
	On Error Resume Next
	ShowDescription(oEvent.Selected)
End Sub

Sub MapList_disposing(oEvent As Object)
End Sub

Sub PaletteList_itemStateChanged(oEvent As Object)
	On Error Resume Next
	ShowPalette(oEvent.Selected)
	' Another palette replaces the custom classes
	If sDialogClasses <> "" Then
		sDialogClasses = ""
		ShowClassesState()
	End If
End Sub

Sub ShowClassesState()
	Dim aB As Variant, aC As Variant
	On Error Resume Next
	If ReadClasses(sDialogClasses, aB, aC) Then
		oDialog.getControl("lblClasses").getModel().Label = Tr("dlg.classes_custom", UBound(aC) + 1)
	Else
		oDialog.getControl("lblClasses").getModel().Label = Tr("dlg.classes_linear")
	End If
End Sub

' "Classes…" button: classes editor on the data of the chosen values
' column
Sub ClassesButton_actionPerformed(oEvent As Object)
	Dim iPal As Integer, s As String
	On Error GoTo Failed
	iPal = oDialog.getControl("lstPalette").getSelectedItemPos()
	If iPal < 0 Then iPal = 0
	s = ClassesEditor(sDialogClasses, aDialogPalettes(iPal), aDialogData, _
		oDialog.getControl("lstCode").getSelectedItemPos(), _
		oDialog.getControl("lstValue").getSelectedItemPos(), _
		oDialog.getControl("lstLabel").getSelectedItemPos(), oDialog.getPeer())
	If s <> Chr(0) Then
		sDialogClasses = s
		ShowClassesState()
	End If
	Exit Sub
Failed:
	WriteLog("ERROR opening the classes editor " & Err & " (line " & Erl & "): " & Error$)
End Sub

Sub ClassesButton_disposing(oEvent As Object)
End Sub

' ---------------------------------------------------------------------------
' Classes editor: number of classes, color bar scaled to the
' values, whose separators can be dragged (value shown below),
' number of zones per class, one tick per value under the bar, sorted list
' of the values. Classes are adjoining and cover all the values.
' Returns the classes text ("": linear distribution of the palette),
' or Chr(0) if canceled.
' ---------------------------------------------------------------------------
Function ClassesEditor(sClasses As String, aPalette As Variant, aData As Variant, iCode As Integer, _
		iVal As Integer, iLab As Integer, oParent As Object) As String
	Dim sLinear As String, nRes As Integer
	On Error GoTo Failed
	ClassesEditor = Chr(0)
	sLinear = BuildEditor(sClasses, aPalette, aData, iCode, iVal, iLab, oParent)
	nRes = oEditor.execute()
	If nRes = 1 Then
		ClassesEditor = ClassesText(aEdBounds, aEdColors, nEdClasses)
		' Same as the linear distribution: no custom classes
		If ClassesEditor = sLinear Then ClassesEditor = ""
		WriteLog("Classes: " & IIf(ClassesEditor = "", "linear", ClassesEditor))
	End If
	oEditor.dispose()
	oEditor = Nothing
	Exit Function
Failed:
	WriteLog("ERROR in the classes editor " & Err & " (line " & Erl & "): " & Error$)
	On Error Resume Next
	If Not IsNull(oEditor) Then oEditor.dispose()
	oEditor = Nothing
	ClassesEditor = Chr(0)
End Function

' Builds the editor (oEditor) and its state; returns the classes text
' of the linear distribution
Function BuildEditor(sClasses As String, aPalette As Variant, aData As Variant, iCode As Integer, _
		iVal As Integer, iLab As Integer, oParent As Object) As String
	Dim oModel As Object, i As Long, k As Integer, dMin As Double, dMax As Double, aValueTexts As Variant
	Dim oButtonsListener As Object, oTextsListener As Object, oMouseListener As Object, oMoveListener As Object
	Dim aBL As Variant, aCL As Variant, nL As Integer, nStep As Long, nTicks As Long
	aEdPalette = aPalette
	SortedValues(aData, iCode, iVal, iLab, aEdValues, aValueTexts)
	If UBound(aEdValues) >= 0 Then
		dMin = aEdValues(0)
		dMax = aEdValues(UBound(aEdValues))
	End If
	LinearClasses(aPalette, dMin, dMax, 0, aBL, aCL, nL)
	BuildEditor = ClassesText(aBL, aCL, nL)
	If ReadClasses(sClasses, aEdBounds, aEdColors) Then
		nEdClasses = UBound(aEdColors) + 1
	Else
		aEdBounds = aBL
		aEdColors = aCL
		nEdClasses = nL
	End If
	' The classes cover all the values
	If aEdBounds(0) > dMin Then aEdBounds(0) = aBL(0)
	If aEdBounds(nEdClasses) < dMax Then aEdBounds(nEdClasses) = aBL(nL)
	dBarMin = aEdBounds(0)
	dBarMax = aEdBounds(nEdClasses)
	iEdBound = -1
	bEdDragging = False
	iEdColor = -1
	nEdShown = -1

	oModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
	oModel.Title = Tr("ed.title")
	oModel.Width = 420
	oModel.Height = 196
	AddControl(oModel, "FixedText", "lblInfo", 6, 6, 320, 10, Array("Label"), _
		Array(Tr("ed.info", EdNumber(dMin), EdNumber(dMax), UBound(aEdValues) + 1)))
	AddControl(oModel, "FixedText", "lblCount", 6, 22, 76, 10, Array("Label"), Array(Tr("ed.count")))
	AddControl(oModel, "NumericField", "numClasses", 84, 20, 34, 14, _
		Array("Value", "ValueMin", "ValueMax", "ValueStep", "DecimalAccuracy", "Spin", "StrictFormat", "Repeat"), _
		Array(CDbl(nEdClasses), 2#, CDbl(MAX_CLASSES), 1#, 0, True, True, True))
	AddControl(oModel, "Button", "btnLinear", 124, 20, 100, 14, _
		Array("Label", "PushButtonType", "HelpText"), Array(Tr("ed.linear"), 0, Tr("ed.linear_help")))
	' Bar: one box per class, a handle above each inner bound,
	' the value of each bound below, and a line under the
	' bar (pixel reference). No control overlaps another:
	' LibreOffice does not guarantee the drawing order of overlapping controls.
	AddControl(oModel, "FixedText", "barBase", BAR_X, BAR_Y + BAR_H, BAR_W, 1, _
		Array("BackgroundColor"), Array(RGB(160, 160, 160)))
	For k = 0 To MAX_CLASSES - 1
		AddControl(oModel, "FixedText", "seg" & k, BAR_X, BAR_Y, 1, BAR_H, _
			Array("Align", "VerticalAlign", "HelpText"), _
			Array(1, com.sun.star.style.VerticalAlignment.MIDDLE, Tr("ed.color_help")))
	Next k
	For k = 1 To MAX_CLASSES - 1
		AddControl(oModel, "FixedText", "pgn" & k, BAR_X, BAR_Y - 9, 5, 8, _
			Array("BackgroundColor", "HelpText"), Array(RGB(40, 40, 40), Tr("ed.drag_help")))
	Next k
	For k = 0 To MAX_CLASSES
		AddControl(oModel, "FixedText", "lblBound" & k, BAR_X, BAR_Y + BAR_H + 4, 26, 10, _
			Array("Align"), Array(1))
	Next k
	' One tick per value (at most MAX_TICKS): where the data are
	nStep = (UBound(aEdValues) + 1) \ MAX_TICKS + 1
	For i = 0 To UBound(aEdValues) Step nStep
		AddControl(oModel, "FixedText", "tick" & nTicks, ValueX(aEdValues(i)), _
			BAR_Y + BAR_H + 17, 1, 5, Array("BackgroundColor"), Array(RGB(96, 96, 96)))
		nTicks = nTicks + 1
	Next i
	AddControl(oModel, "FixedText", "lblHelp", 6, BAR_Y + BAR_H + 27, 320, 10, _
		Array("Label", "TextColor"), Array(Tr("ed.help"), RGB(96, 96, 96)))
	AddControl(oModel, "FixedText", "lblBound", 6, 118, 100, 10, Array("Label"), Array(Tr("ed.bound")))
	AddControl(oModel, "Edit", "txtBound", 108, 116, 60, 14, Array("Align", "Enabled"), Array(2, False))
	AddControl(oModel, "FixedText", "lblError", 6, 136, 320, 10, _
		Array("Label", "TextColor"), Array("", RGB(192, 0, 0)))
	AddControl(oModel, "FixedText", "lblValues", 332, 6, 82, 10, Array("Label"), Array(Tr("ed.values")))
	AddControl(oModel, "ListBox", "lstValues", 332, 18, 82, 150, _
		Array("StringItemList", "ReadOnly"), Array(aValueTexts, True))
	AddControl(oModel, "Button", "btnOkEd", 308, 176, 50, 14, _
		Array("Label", "PushButtonType", "DefaultButton"), Array(Tr("dlg.ok"), 0, True))
	AddControl(oModel, "Button", "btnCancelEd", 364, 176, 50, 14, _
		Array("Label", "PushButtonType"), Array(Tr("dlg.cancel"), 2))

	oEditor = CreateUnoService("com.sun.star.awt.UnoControlDialog")
	oEditor.setModel(oModel)
	oButtonsListener = CreateUnoListener("EditorButton_", "com.sun.star.awt.XActionListener")
	oTextsListener = CreateUnoListener("EditorText_", "com.sun.star.awt.XTextListener")
	oMouseListener = CreateUnoListener("Bar_", "com.sun.star.awt.XMouseListener")
	oMoveListener = CreateUnoListener("BarMove_", "com.sun.star.awt.XMouseMotionListener")
	oEditor.getControl("btnLinear").addActionListener(oButtonsListener)
	oEditor.getControl("btnOkEd").addActionListener(oButtonsListener)
	oEditor.getControl("numClasses").addTextListener(oTextsListener)
	oEditor.getControl("txtBound").addTextListener(oTextsListener)
	' Mouse: on the boxes and the separators; the movement is followed over
	' the whole bar (the pointer quickly leaves a narrow separator)
	oEditor.getControl("barBase").addMouseMotionListener(oMoveListener)
	oEditor.getControl("barBase").addMouseListener(oMouseListener)
	For k = 0 To MAX_CLASSES - 1
		oEditor.getControl("seg" & k).addMouseListener(oMouseListener)
		oEditor.getControl("seg" & k).addMouseMotionListener(oMoveListener)
		If k > 0 Then
			oEditor.getControl("pgn" & k).addMouseListener(oMouseListener)
			oEditor.getControl("pgn" & k).addMouseMotionListener(oMoveListener)
		End If
	Next k
	oEditor.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), oParent)
	BarPointers()
	ShowBar()
End Function

' Numeric values of column iVal, sorted, and their texts for the
' list: "389 - Ain (01)"
Sub SortedValues(aData As Variant, iCode As Integer, iVal As Integer, iLab As Integer, _
		aValues As Variant, aTexts As Variant)
	Dim i As Long, n As Long, d As Double, sCode As String, sLab As String
	Dim aV() As Double, aT() As String
	On Error GoTo Finish
	For i = 0 To UBound(aData)
		If UBound(aData(i)) >= iVal And UBound(aData(i)) >= iCode Then
			sCode = NormalizeCode(aData(i)(iCode))
			If sCode <> "" And Not IsBlank(aData(i)(iVal)) Then
				If ToNumber(aData(i)(iVal), d) Then
					sLab = ""
					If iLab >= 0 And UBound(aData(i)) >= iLab Then
						If Not IsBlank(aData(i)(iLab)) Then sLab = Trim(CStr(aData(i)(iLab)))
					End If
					ReDim Preserve aV(n) As Double
					ReDim Preserve aT(n) As String
					aV(n) = d
					aT(n) = IIf(sLab = "", sCode, sLab & " (" & sCode & ")")
					n = n + 1
				End If
			End If
		End If
	Next i
Finish:
	SortValues(aV, aT)
	For i = 0 To n - 1
		aT(i) = EdNumber(aV(i)) & " - " & aT(i)
	Next i
	aValues = aV
	aTexts = aT
End Sub

' Shell sort of values, associated texts moved with them
Sub SortValues(aV As Variant, aT As Variant)
	Dim nGap As Long, i As Long, j As Long, d As Double, s As String
	nGap = (UBound(aV) + 1) \ 2
	Do While nGap > 0
		For i = nGap To UBound(aV)
			d = aV(i) : s = aT(i)
			j = i
			Do While j >= nGap
				If aV(j - nGap) <= d Then Exit Do
				aV(j) = aV(j - nGap) : aT(j) = aT(j - nGap)
				j = j - nGap
			Loop
			aV(j) = d : aT(j) = s
		Next i
		nGap = nGap \ 2
	Loop
End Sub

' Number shown in the editor (decimal separator of the language)
Function EdNumber(ByVal d As Double) As String
	' Rounded to the millionth (without Round, missing from StarBasic outside VBA mode)
	EdNumber = CStr(Int(d * 1000000 + 0.5) / 1000000)
End Function

' Updates all the editor rows
' Position (dialog units) of a value on the bar
Function ValueX(ByVal d As Double) As Long
	ValueX = BAR_X + CLng((d - dBarMin) / (dBarMax - dBarMin) * BAR_W)
End Function

' Value under a point of the bar (pixels, relative to the dialog), rounded
Function PixelValue(ByVal xPixel As Long) As Double
	' barBase: line under the bar, same x and same width
	Dim oRect As Object, dStep As Double, d As Double
	oRect = oEditor.getControl("barBase").getPosSize()
	d = dBarMin + (xPixel - oRect.X) / oRect.Width * (dBarMax - dBarMin)
	dStep = RoundingStep(dBarMax - dBarMin)
	PixelValue = Int(d / dStep + 0.5) * dStep
End Function

' Mouse pointers: double arrow on the separators, hand on the boxes
Sub BarPointers()
	Dim oArrow As Object, oMain As Object, k As Integer
	On Error Resume Next
	oArrow = CreateUnoService("com.sun.star.awt.Pointer")
	oArrow.setType(com.sun.star.awt.SystemPointer.HSIZEBAR)
	oMain = CreateUnoService("com.sun.star.awt.Pointer")
	oMain.setType(com.sun.star.awt.SystemPointer.REFHAND)
	For k = 0 To MAX_CLASSES - 1
		oEditor.getControl("seg" & k).getPeer().setPointer(oMain)
		If k > 0 Then oEditor.getControl("pgn" & k).getPeer().setPointer(oArrow)
	Next k
End Sub

' Updates the bar: boxes (color, number of zones), separators,
' bound values, number of classes and selected bound. With
' bKeepInput, the fields are not rewritten (the change comes from them).
Sub ShowBar(Optional bKeepInput As Variant)
	Dim k As Integer, x1 As Long, x2 As Long, aN As Variant, oModel As Object, bVisible As Boolean
	bEdUpdating = True
	aN = CountClasses()
	For k = 0 To MAX_CLASSES - 1
		oModel = oEditor.getControl("seg" & k).getModel()
		bVisible = (k < nEdClasses)
		If nEdShown <> nEdClasses Then oEditor.getControl("seg" & k).setVisible(bVisible)
		If bVisible Then
			x1 = ValueX(aEdBounds(k))
			x2 = ValueX(aEdBounds(k + 1))
			oModel.PositionX = x1
			oModel.Width = IIf(x2 - x1 < 1, 1, x2 - x1)
			oModel.BackgroundColor = aEdColors(k)
			oModel.TextColor = ContrastTextColor(aEdColors(k))
			oModel.Label = IIf(x2 - x1 >= 8, CStr(aN(k)), "")
		End If
		If k > 0 Then
			If nEdShown <> nEdClasses Then oEditor.getControl("pgn" & k).setVisible(k < nEdClasses)
			If k < nEdClasses Then
				oEditor.getControl("pgn" & k).getModel().PositionX = ValueX(aEdBounds(k)) - 2
				oEditor.getControl("pgn" & k).getModel().BackgroundColor = _
					IIf(k = iEdBound, RGB(220, 0, 0), RGB(40, 40, 40))
			End If
		End If
	Next k
	For k = 0 To MAX_CLASSES
		oModel = oEditor.getControl("lblBound" & k).getModel()
		If nEdShown <> nEdClasses Then oEditor.getControl("lblBound" & k).setVisible(k <= nEdClasses)
		If k <= nEdClasses Then
			x1 = ValueX(aEdBounds(k)) - 13
			If x1 < 0 Then x1 = 0
			oModel.PositionX = x1
			oModel.Label = EdNumber(aEdBounds(k))
			oModel.TextColor = IIf(k = iEdBound, RGB(220, 0, 0), RGB(0, 0, 0))
		End If
	Next k
	nEdShown = nEdClasses
	' Compared with the text: the field value may lag behind what was typed
	If Not IsMissing(bKeepInput) Then
	ElseIf Val(oEditor.getControl("numClasses").getText()) <> nEdClasses Then
		oEditor.getControl("numClasses").setValue(nEdClasses)
	End If
	oEditor.getControl("txtBound").getModel().Enabled = (iEdBound > 0)
	' No IIf here: it evaluates both branches, hence aEdBounds(-1)
	If Not IsMissing(bKeepInput) Then
	ElseIf iEdBound > 0 Then
		oEditor.getControl("txtBound").setText(EdNumber(aEdBounds(iEdBound)))
	Else
		oEditor.getControl("txtBound").setText("")
	End If
	bEdUpdating = False
End Sub

' Number of zones per class, and outside the classes
' Number of zones per class
Function CountClasses() As Variant
	Dim aN() As Long, i As Long, k As Integer
	ReDim aN(nEdClasses - 1) As Long
	For i = 0 To UBound(aEdValues)
		k = ClassOf(aEdBounds, aEdValues(i))
		If k >= 0 Then aN(k) = aN(k) + 1
	Next i
	CountClasses = aN
End Function

' Mouse on the bar: press on a separator (start of the drag) or on
' a box (its color is chosen on release)
Sub Bar_mousePressed(oEvent As Object)
	Dim sName As String
	On Error Resume Next
	sName = oEvent.Source.getModel().Name
	iEdColor = -1
	If Left(sName, 3) = "pgn" Then
		iEdBound = CInt(Mid(sName, 4))
		bEdDragging = True
		oEditor.getControl("lblError").getModel().Label = ""
		ShowBar()
	ElseIf Left(sName, 3) = "seg" Then
		iEdColor = CInt(Mid(sName, 4))
	End If
End Sub

Sub Bar_mouseReleased(oEvent As Object)
	Dim k As Integer, nColor As Long
	On Error Resume Next
	If bEdDragging Then
		bEdDragging = False
		ShowBar()
	ElseIf iEdColor >= 0 Then
		k = iEdColor
		iEdColor = -1
		nColor = PickColor(aEdColors(k), oEditor.getPeer())
		If nColor >= 0 Then
			aEdColors(k) = nColor
			ShowBar()
		End If
	End If
End Sub

' Dragging a separator: the bound follows the mouse, between its neighbors
Sub BarMove_mouseDragged(oEvent As Object)
	Dim d As Double, dStep As Double, k As Integer
	On Error Resume Next
	If Not bEdDragging Or iEdBound <= 0 Then Exit Sub
	k = iEdBound
	d = PixelValue(oEvent.Source.getPosSize().X + oEvent.X)
	dStep = RoundingStep(dBarMax - dBarMin)
	If d < aEdBounds(k - 1) + dStep Then d = aEdBounds(k - 1) + dStep
	If d > aEdBounds(k + 1) - dStep Then d = aEdBounds(k + 1) - dStep
	If d <> aEdBounds(k) Then
		aEdBounds(k) = d
		ShowBar()
	End If
End Sub

Sub BarMove_mouseMoved(oEvent As Object)
End Sub

Sub BarMove_disposing(oEvent As Object)
End Sub

Sub Bar_mouseEntered(oEvent As Object)
End Sub

Sub Bar_mouseExited(oEvent As Object)
End Sub

Sub Bar_disposing(oEvent As Object)
End Sub

' Bound typed in: "To" of row i (= "From" of the next one), or
' "From" of the 1st row
' Number of classes (linear distribution, palette colors) or typed
' value of the selected bound
Sub EditorText_textChanged(oEvent As Object)
	Dim sName As String, d As Double, n As Integer, k As Integer
	On Error Resume Next
	If bEdUpdating Then Exit Sub
	sName = oEvent.Source.getModel().Name
	oEditor.getControl("lblError").getModel().Label = ""
	If sName = "numClasses" Then
		' Text and not value: while typing or when clicking an arrow, the
		' value is not always up to date yet
		n = CInt(Val(oEditor.getControl("numClasses").getText()))
		If n < 2 Or n > MAX_CLASSES Or n = nEdClasses Then Exit Sub
		LinearClasses(aEdPalette, EdMin(), EdMax(), n, aEdBounds, aEdColors, nEdClasses)
		iEdBound = -1
		ShowBar(True)
		oEditor.getControl("txtBound").setText("")
	ElseIf sName = "txtBound" And iEdBound > 0 Then
		k = iEdBound
		If Not ToNumber(oEvent.Source.getText(), d) Then
			oEditor.getControl("lblError").getModel().Label = Tr("ed.not_number")
		ElseIf d <= aEdBounds(k - 1) Or d >= aEdBounds(k + 1) Then
			oEditor.getControl("lblError").getModel().Label = _
				Tr("ed.between", EdNumber(aEdBounds(k - 1)), EdNumber(aEdBounds(k + 1)))
		Else
			aEdBounds(k) = d
			ShowBar(True)
		End If
	End If
End Sub

Sub EditorText_disposing(oEvent As Object)
End Sub

Sub EditorButton_actionPerformed(oEvent As Object)
	Dim sName As String, k As Integer
	On Error GoTo Failed
	sName = oEvent.Source.getModel().Name
	If sName = "btnOkEd" Then
		For k = 1 To nEdClasses
			If aEdBounds(k) <= aEdBounds(k - 1) Then
				oEditor.getControl("lblError").getModel().Label = Tr("ed.not_increasing", k)
				Exit Sub
			End If
		Next k
		oEditor.endDialog(1)
	ElseIf sName = "btnLinear" Then
		' Equal-interval bounds; number of classes and colors kept
		LinearBounds(EdMin(), EdMax(), nEdClasses, aEdBounds)
		oEditor.getControl("lblError").getModel().Label = ""
		iEdBound = -1
		ShowBar()
	End If
	Exit Sub
Failed:
	WriteLog("ERROR in the classes editor " & Err & " (line " & Erl & "): " & Error$)
End Sub

Sub EditorButton_disposing(oEvent As Object)
End Sub

Function EdMin() As Double
	If UBound(aEdValues) >= 0 Then EdMin = aEdValues(0)
End Function

Function EdMax() As Double
	If UBound(aEdValues) >= 0 Then EdMax = aEdValues(UBound(aEdValues))
End Function

' LibreOffice color picker; -1 if canceled
Function PickColor(nColor As Long, oParent As Object) As Long
	Dim oSelector As Object, aRes As Variant, k As Integer
	Dim aProps(0) As New com.sun.star.beans.PropertyValue
	On Error GoTo Failed
	PickColor = -1
	oSelector = CreateUnoServiceWithArguments("com.sun.star.cui.ColorPicker", Array(oParent))
	aProps(0).Name = "Color"
	aProps(0).Value = nColor
	oSelector.setPropertyValues(aProps())
	If oSelector.execute() = 1 Then
		aRes = oSelector.getPropertyValues()
		For k = 0 To UBound(aRes)
			If aRes(k).Name = "Color" Then PickColor = aRes(k).Value
		Next k
	End If
	Exit Function
Failed:
	WriteLog("ERROR in the color picker: " & Error$)
	PickColor = -1
End Function

Sub PaletteList_disposing(oEvent As Object)
End Sub

' "Values", "labels" and "legend" check boxes: the settings that
' depend on them can only be changed if the box is checked
Sub TextChecks_itemStateChanged(oEvent As Object)
	Dim sSuffix As Variant, sCtrl As Variant, bActive As Boolean
	On Error Resume Next
	For Each sSuffix In Array("Values", "Labels")
		bActive = (oDialog.getControl("chk" & sSuffix).getState() = 1)
		For Each sCtrl In Array("cboFont", "cboSize", "btnBold", "btnItalic", "btnUnderline", "chkOutline")
			oDialog.getControl(sCtrl & sSuffix).getModel().Enabled = bActive
		Next sCtrl
	Next sSuffix
	oDialog.getControl("chkLegendValues").getModel().Enabled = (oDialog.getControl("chkLegend").getState() = 1)
End Sub

Sub TextChecks_disposing(oEvent As Object)
End Sub

Sub OtherFile_actionPerformed(oEvent As Object)
	On Error Resume Next
	oDialog.endDialog(3)
End Sub

Sub OtherFile_disposing(oEvent As Object)
End Sub

' ---------------------------------------------------------------------------
' Insertion of the colored map into the Calc sheet, to the right of the data
' ---------------------------------------------------------------------------
Function InsertIntoCalc(oDoc As Object, oDocCalc As Object, oRange As Object, sName As String, _
		bImage As Boolean) As Object
	Dim oSheet As Object, oCell As Object, oPageDraw As Object, oPageCalc As Object
	Dim oShp As Object, nShapes As Long, a As Variant, oBlock As Object
	Dim oSize As New com.sun.star.awt.Size
	On Error GoTo Failed
	InsertIntoCalc = Nothing
	bImage = False
	oSheet = oRange.getSpreadsheet()
	' To the right of the data block next to the selection, after an empty column
	oBlock = oSheet.createCursorByRange(oRange)
	oBlock.collapseToCurrentRegion()
	a = oBlock.getRangeAddress()
	If a.EndColumn < oRange.getRangeAddress().EndColumn Then a = oRange.getRangeAddress()
	oCell = oSheet.getCellByPosition(a.EndColumn + 2, oRange.getRangeAddress().StartRow)
	oPageDraw = oDoc.getDrawPages().getByIndex(0)
	oPageCalc = oSheet.getDrawPage()
	nShapes = oPageDraw.getCount()

	If nShapes <= CALC_SHAPES_LIMIT Then oShp = CopyShapes(oPageDraw, oDocCalc, oPageCalc)
	If IsNull(oShp) Then
		oShp = InsertImage(oPageDraw, oDocCalc, oPageCalc)
		bImage = True
	End If
	If IsNull(oShp) Then
		Warn(Tr("msg.insert_failed"))
		Exit Function
	End If

	' Fixed width, proportions kept, top left corner on the cell
	oSize.Width = CALC_MAP_WIDTH
	oSize.Height = CALC_MAP_WIDTH * oShp.Size.Height / oShp.Size.Width
	oShp.setSize(oSize)
	oShp.setPosition(oCell.Position)
	oShp.Name = FreeShapeName(oPageCalc, sName)
	oShp.Description = sName
	On Error Resume Next
	oShp.Anchor = oCell
	oDocCalc.getCurrentController().select(oShp)
	On Error GoTo Failed
	WriteLog("Map inserted into Calc as " & CellName(oCell) & " (" _
		& IIf(bImage, "image", nShapes & " shapes") & ")")
	If bImage Then Inform(Tr("msg.inserted_as_image", nShapes))
	InsertIntoCalc = oShp
	Exit Function
Failed:
	WriteLog("ERROR inserting into Calc " & Err & " (line " & Erl & "): " & Error$)
	Warn(Tr("msg.insert_failed"))
End Function

' Recreates the map zones in the sheet and groups them. Each zone
' keeps geometry, colors, name and alternative text. Nothing if impossible
' (the shapes already created are then removed).
Function CopyShapes(oPageDraw As Object, oDocCalc As Object, oPageCalc As Object) As Object
	Dim oColl As Object, oSource As Object, oCopy As Object, sType As String
	Dim i As Long, nSkipped As Long
	On Error GoTo Failed
	CopyShapes = Nothing
	oColl = CreateUnoService("com.sun.star.drawing.ShapeCollection")
	For i = 0 To oPageDraw.getCount() - 1
		oSource = oPageDraw.getByIndex(i)
		sType = oSource.getShapeType()
		If sType = "com.sun.star.drawing.PolyPolygonShape" Or sType = "com.sun.star.drawing.ClosedBezierShape" Then
			oCopy = oDocCalc.createInstance(sType)
			oPageCalc.add(oCopy)
			oColl.add(oCopy)
			If sType = "com.sun.star.drawing.PolyPolygonShape" Then
				oCopy.PolyPolygon = oSource.PolyPolygon
			Else
				oCopy.PolyPolygonBezier = oSource.PolyPolygonBezier
			End If
			oCopy.FillStyle = oSource.FillStyle
			oCopy.FillColor = oSource.FillColor
			oCopy.LineStyle = oSource.LineStyle
			oCopy.LineColor = oSource.LineColor
			oCopy.LineWidth = oSource.LineWidth
			oCopy.Name = oSource.Name
			oCopy.Description = oSource.Description
			If oSource.getString() <> "" Then CopyText(oSource, oCopy, CALC_MAP_WIDTH / oPageDraw.Width)
		Else
			nSkipped = nSkipped + 1
		End If
	Next i
	If nSkipped > 0 Then WriteLog(nSkipped & " non-geographic shape(s) not copied")
	If oColl.getCount() = 0 Then Exit Function
	If oColl.getCount() = 1 Then
		CopyShapes = oColl.getByIndex(0)
	Else
		CopyShapes = oPageCalc.group(oColl)
	End If
	Exit Function
Failed:
	WriteLog("ERROR copying the shapes: " & Error$ & " ; falling back to an image")
	On Error Resume Next
	For i = 0 To oColl.getCount() - 1
		oPageCalc.remove(oColl.getByIndex(i))
	Next i
	CopyShapes = Nothing
End Function

' Exports the map page as PNG and inserts it into the sheet.
' Nothing if impossible.
Function InsertImage(oPageDraw As Object, oDocCalc As Object, oPageCalc As Object) As Object
	Dim oGraphic As Object, oImage As Object
	Dim oSize As New com.sun.star.awt.Size
	On Error GoTo Failed
	InsertImage = Nothing
	oGraphic = PageImage(oPageDraw)
	If IsNull(oGraphic) Then Exit Function
	oImage = oDocCalc.createInstance("com.sun.star.drawing.GraphicObjectShape")
	oPageCalc.add(oImage)
	oImage.Graphic = oGraphic
	oSize.Width = oPageDraw.Width
	oSize.Height = oPageDraw.Height
	oImage.setSize(oSize)
	InsertImage = oImage
	Exit Function
Failed:
	WriteLog("ERROR inserting the image: " & Error$)
	InsertImage = Nothing
End Function

' Draw page -> PNG image (Graphic); Nothing if impossible
Function PageImage(oPageDraw As Object) As Object
	Dim oFilter As Object, sUrl As String
	Dim aDataArray(1) As New com.sun.star.beans.PropertyValue
	Dim aArgs(2) As New com.sun.star.beans.PropertyValue
	Dim aSource(0) As New com.sun.star.beans.PropertyValue
	On Error GoTo Failed
	PageImage = Nothing
	sUrl = MapsFolder() & IMAGE_NAME
	DeleteFile(sUrl)
	aDataArray(0).Name = "PixelWidth"
	aDataArray(0).Value = IMAGE_WIDTH_PX
	aDataArray(1).Name = "PixelHeight"
	aDataArray(1).Value = CLng(IMAGE_WIDTH_PX * oPageDraw.Height / oPageDraw.Width)
	aArgs(0).Name = "URL"
	aArgs(0).Value = sUrl
	aArgs(1).Name = "MediaType"
	aArgs(1).Value = "image/png"
	aArgs(2).Name = "FilterData"
	aArgs(2).Value = aDataArray()
	oFilter = CreateUnoService("com.sun.star.drawing.GraphicExportFilter")
	oFilter.setSourceDocument(oPageDraw)
	oFilter.filter(aArgs())

	aSource(0).Name = "URL"
	aSource(0).Value = sUrl
	PageImage = CreateUnoService("com.sun.star.graphic.GraphicProvider").queryGraphic(aSource())
	Exit Function
Failed:
	WriteLog("ERROR exporting the image: " & Error$)
	PageImage = Nothing
End Function

' Shape name not used on the page: "name", otherwise "name (2)", "name (3)"…
Function FreeShapeName(oPage As Object, sName As String) As String
	Dim n As Integer, s As String
	s = sName
	n = 1
	Do While Not IsNull(FindShape(oPage, s))
		n = n + 1
		s = sName & " (" & n & ")"
	Loop
	FreeShapeName = s
End Function

' Top-level shape with this name; Nothing if missing
Function FindShape(oPage As Object, sName As String) As Object
	Dim i As Long
	On Error GoTo Failed
	FindShape = Nothing
	For i = 0 To oPage.getCount() - 1
		If oPage.getByIndex(i).Name = sName Then
			FindShape = oPage.getByIndex(i)
			Exit Function
		End If
	Next i
Failed:
End Function

' ---------------------------------------------------------------------------
' Automatic update of the maps inserted into Calc.
' A link = a named range CalcMap_N (the 2 selected columns,
' which follows row insertions) + a document property of the same name:
' "1|shape name|palette id|base map URL|image (0/1)|options"
' (LinkFields; options in O_* order, separated by |; links from
' previous versions have fewer: default values).
' ---------------------------------------------------------------------------
Function CreateLink(oDocCalc As Object, oSelection As Object, oShp As Object, sPalette As String, _
		sMapUrl As String, bImage As Boolean, aOptions As Variant) As String
	Dim oRanges As Object, sName As String, n As Integer, oProps As Object
	Dim oOrigin As New com.sun.star.table.CellAddress
	On Error GoTo Failed
	CreateLink = ""
	oRanges = oDocCalc.NamedRanges
	n = 1
	Do While oRanges.hasByName(LINK_PREFIX & n)
		n = n + 1
	Loop
	sName = LINK_PREFIX & n
	oRanges.addNewByName(sName, oSelection.AbsoluteName, oOrigin, 0)
	oProps = oDocCalc.getDocumentProperties().getUserDefinedProperties()
	oProps.addProperty(sName, com.sun.star.beans.PropertyAttribute.REMOVABLE, _
		LinkFields(oShp.Name, sPalette, sMapUrl, bImage, aOptions))
	WriteLog("Map linked to the data " & oSelection.AbsoluteName & " (" & sName & "): automatic update")
	WatchLink(oDocCalc, sName)
	CreateLink = sName
	Exit Function
Failed:
	WriteLog("ERROR: the map will not be updated automatically (" & Error$ & ")")
End Function

Function LinkFields(sShape As String, sPalette As String, sMapUrl As String, bImage As Boolean, _
		aOptions As Variant) As String
	Dim s As String, k As Integer
	s = "1|" & sShape & "|" & sPalette & "|" & sMapUrl & "|" & IIf(bImage, "1", "0")
	For k = 0 To OPTION_COUNT - 1
		s = s & "|" & OptionToText(k, aOptions(k))
	Next k
	LinkFields = s
End Function

' Options of a link (missing fields: default values)
Function LinkOptions(aFields As Variant) As Variant
	Dim aOptions As Variant, k As Integer
	aOptions = DefaultOptions()
	For k = 0 To OPTION_COUNT - 1
		If UBound(aFields) >= 5 + k Then aOptions(k) = OptionFromText(k, aFields(5 + k))
	Next k
	LinkOptions = aOptions
End Function

' Listens to the changes of a link's range (only once per link)
Sub WatchLink(oDocCalc As Object, sName As String)
	Dim oRange As Object, oListener As Object, i As Long
	On Error GoTo Failed
	For i = 0 To nLinks - 1
		If EqualUnoObjects(aLinks(i)(0), oDocCalc) And aLinks(i)(3) = sName Then Exit Sub
	Next i
	oRange = oDocCalc.NamedRanges.getByName(sName).getReferredCells()
	If IsNull(oRange) Then Exit Sub
	oListener = CreateUnoListener("MapData_", "com.sun.star.util.XModifyListener")
	oRange.addModifyListener(oListener)
	ReDim Preserve aLinks(nLinks)
	aLinks(nLinks) = Array(oDocCalc, oRange, oListener, sName, Empty)
	nLinks = nLinks + 1
	InterceptMenu(oDocCalc)
	Exit Sub
Failed:
	WriteLog("ERROR watching " & sName & ": " & Error$)
End Sub

' "Open Document" event (Events.xcu): reconnects the linked maps
' of a spreadsheet. Silent and fast for any other document.
Sub ReconnectMaps(Optional oEvent As Variant)
	Dim oDoc As Object, oProps As Object, aNames As Variant, i As Long
	On Error GoTo Finish
	If IsMissing(oEvent) Then
		oDoc = ThisComponent
	Else
		oDoc = oEvent.Source
	End If
	If IsNull(oDoc) Then Exit Sub
	If Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then Exit Sub
	oProps = oDoc.getDocumentProperties().getUserDefinedProperties()
	aNames = oProps.getPropertySetInfo().getProperties()
	For i = 0 To UBound(aNames)
		If Left(aNames(i).Name, Len(LINK_PREFIX)) = LINK_PREFIX _
				Or Left(aNames(i).Name, Len(OLD_LINK_PREFIX)) = OLD_LINK_PREFIX Then
			If oDoc.NamedRanges.hasByName(aNames(i).Name) Then WatchLink(oDoc, aNames(i).Name)
		End If
	Next i
Finish:
End Sub

' Listener: a value of the range has changed
Sub MapData_modified(oEvent As Object)
	Dim i As Long
	On Error GoTo Finish
	If bRecalcRunning Then Exit Sub
	For i = 0 To nLinks - 1
		If EqualUnoObjects(aLinks(i)(1), oEvent.Source) Then
			bRecalcRunning = True
			RecalcLink(aLinks(i)(0), aLinks(i)(3))
			bRecalcRunning = False
			Exit Sub
		End If
	Next i
Finish:
	bRecalcRunning = False
End Sub

' Spreadsheet closed: forgets its links
Sub MapData_disposing(oEvent As Object)
	Dim i As Long, k As Long
	On Error Resume Next
	For i = nLinks - 1 To 0 Step -1
		If EqualUnoObjects(aLinks(i)(1), oEvent.Source) Then
			For k = i To nLinks - 2
				aLinks(k) = aLinks(k + 1)
			Next k
			nLinks = nLinks - 1
		End If
	Next i
End Sub

' Recolors the map of a link with the current data. If the map has been
' deleted from the sheet, the link is removed. Without bFull, only the
' zones whose value or label changed since the last update
' are redone, when possible (PartialUpdate).
Sub RecalcLink(oDocCalc As Object, sName As String, Optional bFull As Variant)
	Dim oRange As Object, aFields As Variant, oShp As Object, aData As Variant
	Dim aPalette As Variant, aPalettes As Variant, oUndo As Object, oColl As Object
	Dim oDoc As Object, oPage As Object, vText As Variant, sFields As String, iLink As Long, bDone As Boolean
	Dim aProps(1) As New com.sun.star.beans.PropertyValue
	On Error GoTo Failed
	oRange = oDocCalc.NamedRanges.getByName(sName).getReferredCells()
	sFields = oDocCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sName)
	aFields = Split(sFields, "|")
	If UBound(aFields) < 4 Then Exit Sub
	oShp = FindShape(oRange.getSpreadsheet().getDrawPage(), aFields(1))
	If IsNull(oShp) Then
		ForgetLink(oDocCalc, sName)
		Exit Sub
	End If
	aPalettes = LoadPalettes()
	aPalette = aPalettes(PaletteIndex(aPalettes, aFields(2)))
	oRange = DataRange(oRange)
	If IsNull(oRange) Then Exit Sub
	aData = oRange.getDataArray()
	' Always a text context: a text disabled by the configuration
	' is thus cleared
	vText = TextContext(oDocCalc, oRange, aData, LinkOptions(aFields), oDocCalc)

	' The recoloring does not go into the undo history: Ctrl+Z undoes
	' the user's input, not the color changes
	oUndo = oDocCalc.getUndoManager()
	oUndo.enterHiddenUndoContext()
	oDocCalc.lockControllers()
	iLink = LinkIndex(oDocCalc, sName)
	nLastUpdate = -1
	If IsMissing(bFull) And iLink >= 0 And aFields(4) <> "1" Then
		If oShp.supportsService("com.sun.star.drawing.GroupShape") Then
			bDone = PartialUpdate(oShp, aData, aPalette, vText, aLinks(iLink)(4), sFields)
		End If
	End If
	If bDone Then
	ElseIf aFields(4) = "1" Then
		' Image (detailed map): new rendering from the base map
		aProps(0).Name = "Hidden"
		aProps(0).Value = True
		aProps(1).Name = "AsTemplate"
		aProps(1).Value = True
		oDoc = StarDesktop.loadComponentFromURL(aFields(3), "_blank", 0, aProps())
		vText(T_DOCUMENT) = oDoc
		ColorDocument(oDoc, aData, aPalette, vText)
		oPage = oDoc.getDrawPages().getByIndex(0)
		oShp.Graphic = PageImage(oPage)
		oDoc.close(True)
	ElseIf oShp.supportsService("com.sun.star.drawing.GroupShape") Then
		ColorContainers(Array(oShp), aData, aPalette, vText)
	Else
		oColl = CreateUnoService("com.sun.star.drawing.ShapeCollection")
		oColl.add(oShp)
		ColorContainers(Array(oColl), aData, aPalette, vText)
	End If
	' Memory for the next partial update
	If iLink >= 0 Then RememberMap(iLink, MapMemory(sFields, aData, vText(T_OPTIONS), vText(T_TITLE)))
	oDocCalc.unlockControllers()
	oUndo.leaveUndoContext()
	Exit Sub
Failed:
	WriteLog("ERROR recalculating " & sName & " " & Err & " (line " & Erl & "): " & Error$)
	On Error Resume Next
	If Not IsNull(oDoc) Then oDoc.close(True)
	If Not IsNull(oUndo) Then oUndo.leaveUndoContext()
	If oDocCalc.hasControllersLocked() Then oDocCalc.unlockControllers()
End Sub

' ---------------------------------------------------------------------------
' Configuration of an inserted map: right-click on the map ->
' "Map settings…". The context menu of the spreadsheet window
' is intercepted (once per window) as soon as a linked map in it is
' watched.
' ---------------------------------------------------------------------------
Sub InterceptMenu(oDocCalc As Object)
	Dim oCtrl As Object, oInterceptor As Object, i As Long
	On Error GoTo Failed
	oCtrl = oDocCalc.getCurrentController()
	If IsNull(oCtrl) Then Exit Sub
	For i = 0 To nMenus - 1
		If EqualUnoObjects(aMenus(i)(0), oCtrl) Then Exit Sub
	Next i
	oInterceptor = CreateUnoListener("MapMenu_", "com.sun.star.ui.XContextMenuInterceptor")
	oCtrl.registerContextMenuInterceptor(oInterceptor)
	ReDim Preserve aMenus(nMenus)
	aMenus(nMenus) = Array(oCtrl, oInterceptor)
	nMenus = nMenus + 1
	Exit Sub
Failed:
	WriteLog("ERROR adding the context menu: " & Error$)
End Sub

' Adds "Map settings…" at the top of the menu if the right-click
' is on a linked map (or one of its zones)
Function MapMenu_notifyContextMenuExecute(oEvent As Object) As Variant
	Dim oSel As Object, sLink As String, oMenu As Object, oEntry As Object, oSeparator As Object
	On Error GoTo Skip
	MapMenu_notifyContextMenuExecute = com.sun.star.ui.ContextMenuInterceptorAction.IGNORED
	oSel = oEvent.Selection.getSelection()
	' All the watched maps are examined: the document is not
	' always reachable from the event
	sLink = SelectionLink(Nothing, oSel)
	AppendLog("Context menu: selection " & ImplName(oSel) & ", linked map '" & sLink & "'")
	If sLink = "" Then Exit Function
	' Properties written with setPropertyValue: Basic does not see them as
	' oEntry.Text ("object variable not set", LibreOffice 25.2)
	oMenu = oEvent.ActionTriggerContainer
	oEntry = oMenu.createInstance("com.sun.star.ui.ActionTrigger")
	oEntry.setPropertyValue("Text", Tr("menu.configure"))
	oEntry.setPropertyValue("CommandURL", CONFIGURE_URL)
	' Icon: the one declared for CONFIGURE_URL in Addons.xcu (LibreOffice
	' 25.2 ignores the Image property of the entry)
	oSeparator = oMenu.createInstance("com.sun.star.ui.ActionTriggerSeparator")
	oSeparator.setPropertyValue("SeparatorType", com.sun.star.ui.ActionTriggerSeparatorType.LINE)
	oMenu.insertByIndex(0, oSeparator)
	oMenu.insertByIndex(0, oEntry)
	AppendLog("Menu: entry """ & Tr("menu.configure") & """ added")
	MapMenu_notifyContextMenuExecute = com.sun.star.ui.ContextMenuInterceptorAction.EXECUTE_MODIFIED
	Exit Function
Skip:
	AppendLog("ERROR in the context menu " & Err & " (line " & Erl & "): " & Error$)
	MapMenu_notifyContextMenuExecute = com.sun.star.ui.ContextMenuInterceptorAction.IGNORED
End Function

Function ImplName(o As Object) As String
	On Error GoTo Failed
	ImplName = "(none)"
	If Not IsNull(o) Then ImplName = o.getImplementationName()
	Exit Function
Failed:
	ImplName = "(unknown)"
End Function

' Link name (CalcMap_N) of the selected map, or of a zone of this
' map; "" if the selection is not a linked map. oDocCalc Nothing:
' maps of all documents
Function SelectionLink(oDocCalc As Object, oSel As Object) As String
	Dim oZone As Object, oShp As Object, aFields As Variant, i As Long, k As Long
	On Error GoTo Failed
	SelectionLink = ""
	If IsNull(oSel) Then Exit Function
	If HasUnoInterfaces(oSel, "com.sun.star.drawing.XShape") Then
		oZone = oSel
	ElseIf HasUnoInterfaces(oSel, "com.sun.star.drawing.XShapes") Then
		If oSel.getCount() = 0 Then Exit Function
		oZone = oSel.getByIndex(0)
	Else
		Exit Function
	End If
	For i = 0 To nLinks - 1
		If IsNull(oDocCalc) Or EqualUnoObjects(aLinks(i)(0), oDocCalc) Then
			aFields = Split(aLinks(i)(0).getDocumentProperties().getUserDefinedProperties() _
				.getPropertyValue(aLinks(i)(3)), "|")
			oShp = FindShape(aLinks(i)(1).getSpreadsheet().getDrawPage(), aFields(1))
			If Not IsNull(oShp) Then
				If EqualUnoObjects(oShp, oZone) Then
					SelectionLink = aLinks(i)(3)
					Exit Function
				End If
				If oShp.supportsService("com.sun.star.drawing.GroupShape") Then
					For k = 0 To oShp.getCount() - 1
						If EqualUnoObjects(oShp.getByIndex(k), oZone) Then
							SelectionLink = aLinks(i)(3)
							Exit Function
						End If
					Next k
				End If
			End If
		End If
	Next i
	Exit Function
Failed:
	SelectionLink = ""
End Function

' Map of a link for the configuration dialog: the one from the cached
' catalog (title, description, preview) if its file is there
Function LinkMap(sMapUrl As String) As Variant
	Dim aParts As Variant, sFile As String, aCatalog As Variant, b As Boolean, i As Long
	aParts = Split(sMapUrl, "/")
	sFile = aParts(UBound(aParts))
	LinkMap = NewMap("", FileNameOf(sMapUrl), "", sMapUrl, sFile, 0, "", True)
	On Error GoTo Finish
	aCatalog = LoadCatalog(MapsFolder(), b, True)
	For i = 0 To UBound(aCatalog)
		If aCatalog(i)(C_FILE) = sFile Then LinkMap = aCatalog(i)
	Next i
Finish:
End Function

' "Map settings…" menu: reopens the dialog for the selected
' map (palette, values, labels), then recolors it
Sub ConfigureMap()
	Dim oDocCalc As Object, sName As String, aFields As Variant, aOptions As Variant
	Dim aPalettes As Variant, aMaps As Variant, iChoice As Integer, iPalette As Integer
	Dim bInsert As Boolean, nRes As Integer, sPrefMap As String, sPrefPalette As String
	Dim bPrefInsert As Boolean, aPrefOptions As Variant, aColumns As Variant, oRange As Object
	On Error GoTo Failed
	OpenLog()
	WriteLog("CalcMap " & VERSION & " - map configuration")
	LoadTexts()
	oDocCalc = ThisComponent
	sName = SelectionLink(oDocCalc, oDocCalc.getCurrentController().getSelection())
	If sName = "" Then
		Warn(Tr("msg.select_map"))
		GoTo Finish
	End If
	aFields = Split(oDocCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sName), "|")
	If UBound(aFields) < 4 Then GoTo Finish
	aOptions = LinkOptions(aFields)
	oRange = DataRange(oDocCalc.NamedRanges.getByName(sName).getReferredCells())
	If IsNull(oRange) Then GoTo Finish
	aColumns = ColumnNames(oRange)
	aDialogData = oRange.getDataArray()
	aPalettes = LoadPalettes()
	aMaps = Array(LinkMap(aFields(3)))
	bInsert = True
	nRes = MapDialog(aMaps, aPalettes, "", "", aFields(2), iChoice, iPalette, bInsert, aOptions, _
		aColumns, ParentWindow(oDocCalc), True)
	WriteLog("Configuration of " & sName & ": result " & nRes & ", palette no. " & iPalette & ", " _
		& OptionsText(aOptions))
	If nRes <> 1 Or iPalette < 0 Then GoTo Finish
	oDocCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sName, _
		LinkFields(aFields(1), aPalettes(iPalette)(P_ID), aFields(3), aFields(4) = "1", aOptions))
	' The choices become the default ones for the next map
	' (except the columns, not remembered)
	ReadPreferences(sPrefMap, sPrefPalette, bPrefInsert, aPrefOptions)
	WritePreferences(sPrefMap, aPalettes(iPalette)(P_ID), bPrefInsert, aOptions)
	RecalcLink(oDocCalc, sName)
	GoTo Finish
Failed:
	WriteLog("ERROR configuring the map " & Err & " (line " & Erl & "): " & Error$)
	On Error Resume Next
	Warn(Tr("msg.partial"))
Finish:
	On Error Resume Next
	WriteLog("Done.")
	CloseLog()
End Sub

Function LinkIndex(oDocCalc As Object, sName As String) As Long
	Dim i As Long
	LinkIndex = -1
	For i = 0 To nLinks - 1
		If EqualUnoObjects(aLinks(i)(0), oDocCalc) And aLinks(i)(3) = sName Then
			LinkIndex = i
			Exit Function
		End If
	Next i
End Function

' Memory of a map after its update: Array(link text, values
' (code -> value), labels (code -> label), min, max, title, number of
' values)
Function MapMemory(sFields As String, aData As Variant, aOptions As Variant, sTitle As String) As Variant
	Dim oValues As Object, oLabels As Object, dMin As Double, dMax As Double, n As Long
	oValues = com.sun.star.container.EnumerableMap.create("string", "double")
	oLabels = com.sun.star.container.EnumerableMap.create("string", "string")
	n = ReadData(aData, oValues, dMin, dMax, oLabels, aOptions(O_COL_CODE), aOptions(O_COL_VALUE), _
		aOptions(O_COL_LABEL))
	MapMemory = Array(sFields, oValues, oLabels, dMin, dMax, sTitle, n)
End Function

Sub RememberMap(i As Long, aMemory As Variant)
	Dim a As Variant
	a = aLinks(i)
	aLinks(i) = Array(a(0), a(1), a(2), a(3), aMemory)
End Sub

' Zone by zone update: only redoes the zones whose value or
' label changed (color, alternative text, texts). Returns False if a
' full update is needed: no memory, configuration
' changed, minimum, maximum or title changed (colors or legend of the whole
' map), zone that gains or loses its value ("No data" in the legend),
' or too many zones changed.
Function PartialUpdate(oGroup As Object, aData As Variant, aPalette As Variant, vText As Variant, _
		vMemory As Variant, sFields As String) As Boolean
	Dim aOptions As Variant, aNew As Variant, oChanges As Object, oEnum As Object, sKey As String, n As Long
	Dim aZones() As Object, nZ As Long, i As Long, oZone As Object, sName As String, dVal As Double
	Dim dScale As Double, dValueSize As Double, dLabelSize As Double, sLabel As String, sValue As String
	Dim nColor As Long, bCustom As Boolean, aBounds As Variant, aCustomColors As Variant
	PartialUpdate = False
	If Not IsArray(vMemory) Then Exit Function
	If vMemory(0) <> sFields Then Exit Function
	aOptions = vText(T_OPTIONS)
	aNew = MapMemory(sFields, aData, aOptions, vText(T_TITLE))
	If aNew(6) = 0 Then Exit Function
	If aNew(3) <> vMemory(3) Or aNew(4) <> vMemory(4) Or aNew(5) <> vMemory(5) Then Exit Function

	' Codes whose value or label changed
	oChanges = com.sun.star.container.EnumerableMap.create("string", "boolean")
	If Not CompareValues(vMemory(1), aNew(1), oChanges) Then Exit Function
	If Not CompareValues(aNew(1), vMemory(1), oChanges) Then Exit Function
	CompareLabels(vMemory(2), aNew(2), oChanges)
	CompareLabels(aNew(2), vMemory(2), oChanges)
	oEnum = oChanges.createKeyEnumeration(False)
	Do While oEnum.hasMoreElements()
		oEnum.nextElement()
		n = n + 1
	Loop
	If n > PARTIAL_MAX_ZONES Then Exit Function
	PartialUpdate = True
	nLastUpdate = 0
	If n = 0 Then Exit Function

	' Zones concerned (collected before adding or removing texts)
	For i = 0 To oGroup.getCount() - 1
		oZone = oGroup.getByIndex(i)
		If Not IsAddOn(oZone.Name) Then
			If FoundKey(oChanges, oZone.Name) <> "" Then
				ReDim Preserve aZones(nZ) As Object
				aZones(nZ) = oZone
				nZ = nZ + 1
			End If
		End If
	Next i
	dScale = ZonesWidth(oGroup) / CALC_MAP_WIDTH
	If aOptions(O_VALUES) Then dValueSize = aOptions(O_SIZE_VALUES) * dScale
	If aOptions(O_LABELS) Then dLabelSize = aOptions(O_SIZE_LABELS) * dScale
	bCustom = ReadClasses(aOptions(O_CLASSES), aBounds, aCustomColors)
	For i = 0 To nZ - 1
		oZone = aZones(i)
		sName = oZone.Name
		RemoveZoneTexts(oGroup, sName)
		sLabel = ""
		If dLabelSize > 0 Then sLabel = FindLabel(aNew(2), sName)
		sValue = ""
		If FindValue(aNew(1), sName, dVal, sKey) Then
			nColor = ValueColor(dVal, aNew(3), aNew(4), aPalette, bCustom, aBounds, aCustomColors)
			If nColor < 0 Then nColor = aPalette(P_EMPTY)
			ApplyZone(oZone, nColor, Tr("alt.value", dVal), True)
			If dValueSize > 0 Then sValue = ValueText(vText, dVal)
		Else
			nColor = aPalette(P_EMPTY)
			ApplyZone(oZone, nColor, Tr("alt.none"), True)
		End If
		ZoneText(vText(T_DOCUMENT), oGroup, oZone, sLabel, dLabelSize, sValue, dValueSize, aOptions, nColor)
	Next i
	WriteLog("Partial update: " & nZ & " zone(s)")
	nLastUpdate = nZ
End Function

' Adds to oChanges the codes of oA missing from oB or with a different value.
' False if a code has a value in oA and none in oB (full update
' needed)
Function CompareValues(oA As Object, oB As Object, oChanges As Object) As Boolean
	Dim oEnum As Object, sKey As String
	CompareValues = False
	oEnum = oA.createKeyEnumeration(False)
	Do While oEnum.hasMoreElements()
		sKey = oEnum.nextElement()
		If Not oB.containsKey(sKey) Then Exit Function
		If oB.get(sKey) <> oA.get(sKey) Then oChanges.put(sKey, True)
	Loop
	CompareValues = True
End Function

' Adds to oChanges the codes whose label differs (or is missing) in oB
Sub CompareLabels(oA As Object, oB As Object, oChanges As Object)
	Dim oEnum As Object, sKey As String
	oEnum = oA.createKeyEnumeration(False)
	Do While oEnum.hasMoreElements()
		sKey = oEnum.nextElement()
		If Not oB.containsKey(sKey) Then
			oChanges.put(sKey, True)
		ElseIf oB.get(sKey) <> oA.get(sKey) Then
			oChanges.put(sKey, True)
		End If
	Loop
End Sub

' Removes the texts of a zone (text and outline copies)
Sub RemoveZoneTexts(oContainer As Object, sName As String)
	Dim i As Long
	For i = oContainer.getCount() - 1 To 0 Step -1
		If oContainer.getByIndex(i).Name = TEXT_PREFIX & " " & sName _
			Or oContainer.getByIndex(i).Name = OLD_TEXT_PREFIX & " " & sName Then oContainer.remove(oContainer.getByIndex(i))
	Next i
End Sub

' Removes a link whose map has disappeared (named range, property, listener)
Sub ForgetLink(oDocCalc As Object, sName As String)
	Dim i As Long, k As Long
	On Error Resume Next
	oDocCalc.getDocumentProperties().getUserDefinedProperties().removeProperty(sName)
	oDocCalc.NamedRanges.removeByName(sName)
	For i = nLinks - 1 To 0 Step -1
		If EqualUnoObjects(aLinks(i)(0), oDocCalc) And aLinks(i)(3) = sName Then
			aLinks(i)(1).removeModifyListener(aLinks(i)(2))
			For k = i To nLinks - 2
				aLinks(k) = aLinks(k + 1)
			Next k
			nLinks = nLinks - 1
		End If
	Next i
End Sub

Function CellName(oCell As Object) As String
	On Error GoTo Failed
	CellName = oCell.AbsoluteName
	Exit Function
Failed:
	CellName = "?"
End Function

' File name without extension, from a URL
Function FileNameOf(sUrl As String) As String
	Dim aParts As Variant, s As String
	aParts = Split(ConvertFromURL(sUrl), GetPathSeparator())
	s = aParts(UBound(aParts))
	If LCase(Right(s, 4)) = ".odg" Then s = Left(s, Len(s) - 4)
	FileNameOf = s
End Function

' ---------------------------------------------------------------------------
' Color palettes: palettes.json of the extension, then the user's one
' (cache folder), which adds palettes or replaces
' those with the same id.
' Palette -> Array(id, name, colors(), no-data color, continuous gradient).
' ---------------------------------------------------------------------------
Function LoadPalettes() As Variant
	Dim aRes() As Variant, n As Long, aRead As Variant, sFolder As String
	Dim i As Long, k As Long, bReplaced As Boolean
	On Error GoTo Failed
	sFolder = ExtensionFolder()
	If sFolder <> "" Then aRes = ReadPalettes(sFolder & PALETTES_NAME)
	n = UBound(aRes) + 1
	sFolder = MapsFolder()
	If sFolder <> "" Then
		aRead = ReadPalettes(sFolder & PALETTES_NAME)
		For i = 0 To UBound(aRead)
			bReplaced = False
			For k = 0 To n - 1
				If aRes(k)(P_ID) = aRead(i)(P_ID) Then
					aRes(k) = aRead(i)
					bReplaced = True
				End If
			Next k
			If Not bReplaced Then
				ReDim Preserve aRes(n)
				aRes(n) = aRead(i)
				n = n + 1
			End If
		Next i
	End If
Finish:
	If n = 0 Then
		ReDim aRes(0)
		aRes(0) = NewPalette(DEFAULT_PALETTE, "Blues", _
			Array(RGB(239, 243, 255), RGB(189, 215, 231), RGB(107, 174, 214), RGB(49, 130, 189), RGB(8, 81, 156)), _
			RGB(230, 230, 230), False)
	End If
	WriteLog("Available palettes: " & UBound(aRes) + 1)
	LoadPalettes = aRes
	Exit Function
Failed:
	WriteLog("ERROR loading palettes: " & Error$)
	Resume Finish
End Function

' Reads a palettes file; skips incomplete palettes (log)
Function ReadPalettes(sUrl As String) As Variant
	Dim oRoot As Variant, aList As Variant, oPal As Variant, aShades As Variant
	Dim aRes() As Variant, aColors() As Long, n As Long, i As Long, k As Long
	Dim sId As String, nEmpty As Long, bValid As Boolean, vContinuous As Variant
	On Error GoTo Failed
	ReadPalettes = Array()
	If Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(sUrl) Then Exit Function
	oRoot = JsonRead(ReadText(sUrl))
	aList = JsonElements(JsonField(oRoot, "palettes"))
	For i = 0 To UBound(aList)
		oPal = aList(i)
		sId = JsonText(oPal, "id")
		aShades = JsonElements(JsonField(oPal, "colors"))
		bValid = (sId <> "" And UBound(aShades) >= 1)
		If bValid Then
			ReDim aColors(UBound(aShades)) As Long
			For k = 0 To UBound(aShades)
				aColors(k) = HexColor(aShades(k))
				If aColors(k) < 0 Then bValid = False
			Next k
		End If
		If bValid Then
			nEmpty = HexColor(JsonText(oPal, "nodata"))
			If nEmpty < 0 Then nEmpty = RGB(230, 230, 230)
			vContinuous = JsonField(oPal, "continuous")
			ReDim Preserve aRes(n)
			aRes(n) = NewPalette(sId, PaletteName(oPal, sId), aColors, nEmpty, _
				(VarType(vContinuous) = 11) And (vContinuous = True))
			n = n + 1
		Else
			WriteLog("Palette " & (i + 1) & " of " & ConvertFromURL(sUrl) _
				& " skipped: missing id, fewer than 2 colors or invalid color")
		End If
	Next i
	If n > 0 Then ReadPalettes = aRes
	Exit Function
Failed:
	WriteLog("ERROR reading the palettes " & ConvertFromURL(sUrl) & ": " & Error$)
	If n > 0 Then ReadPalettes = aRes
End Function

Function NewPalette(sId, sName, aColors, nEmpty, bContinuous) As Variant
	NewPalette = Array(sId, sName, aColors, nEmpty, bContinuous)
End Function

' Palette name: plain text, or object {"fr": ..., "en": ...}
Function PaletteName(oPal As Variant, sId As String) As String
	Dim v As Variant, s As String
	v = JsonField(oPal, "name")
	If VarType(v) = 8 Then
		s = v
	Else
		s = JsonText(v, UiLanguage())
		If s = "" Then s = JsonText(v, "en")
	End If
	If s = "" Then s = sId
	PaletteName = s
End Function

' Index of palette sId, otherwise of the default palette, otherwise 0
Function PaletteIndex(aPalettes As Variant, sId As String) As Integer
	Dim i As Integer
	PaletteIndex = 0
	For i = 0 To UBound(aPalettes)
		If aPalettes(i)(P_ID) = DEFAULT_PALETTE Then PaletteIndex = i
	Next i
	For i = 0 To UBound(aPalettes)
		If sId <> "" And aPalettes(i)(P_ID) = sId Then PaletteIndex = i
	Next i
End Function

' "#rrggbb" -> color; -1 if invalid
Function HexColor(v As Variant) As Long
	Dim s As String, k As Integer, nDigit As Integer, nValue As Long
	HexColor = -1
	If VarType(v) <> 8 Then Exit Function
	s = LCase(Trim(v))
	If Left(s, 1) = "#" Then s = Mid(s, 2)
	If Len(s) <> 6 Then Exit Function
	For k = 1 To 6
		nDigit = InStr(1, "0123456789abcdef", Mid(s, k, 1), 0) - 1
		If nDigit < 0 Then Exit Function
		nValue = nValue * 16 + nDigit
	Next k
	HexColor = nValue
End Function

' ---------------------------------------------------------------------------
' Translations: i18n/messages_<language>.properties of the extension. English
' is loaded first, then completed by the interface language.
' ---------------------------------------------------------------------------
Sub LoadTexts()
	Dim sFolder As String
	On Error GoTo Failed
	oTexts = com.sun.star.container.EnumerableMap.create("string", "string")
	sFolder = ExtensionFolder()
	If sFolder = "" Then Exit Sub
	ReadProperties(sFolder & "i18n/messages_en.properties", oTexts)
	If UiLanguage() <> "en" Then ReadProperties(sFolder & "i18n/messages_" & UiLanguage() & ".properties", oTexts)
	WriteLog("Interface language: " & UiLanguage())
	Exit Sub
Failed:
	WriteLog("ERROR loading translations: " & Error$)
End Sub

' Translated text; %1 and %2 are replaced by v1 and v2. Unknown key: the key.
Function Tr(sKey As String, Optional v1 As Variant, Optional v2 As Variant, Optional v3 As Variant) As String
	Dim s As String
	On Error GoTo Failed
	If IsNull(oTexts) Then LoadTexts()
	s = sKey
	If oTexts.containsKey(sKey) Then s = oTexts.get(sKey)
	If Not IsMissing(v1) Then s = Replace(s, "%1", CStr(v1))
	If Not IsMissing(v2) Then s = Replace(s, "%2", CStr(v2))
	If Not IsMissing(v3) Then s = Replace(s, "%3", CStr(v3))
	Tr = s
	Exit Function
Failed:
	Tr = sKey
End Function

' Language code of the interface ("fr", "en"…)
Function UiLanguage() As String
	Dim s As String
	If sLanguage = "" Then
		s = ReadConfiguration("/org.openoffice.Office.Linguistic/General", "UILocale")
		If s = "" Then s = ReadConfiguration("/org.openoffice.Setup/L10N", "ooLocale")
		s = LCase(Left(s, 2))
		If s = "" Then s = "en"
		sLanguage = s
	End If
	UiLanguage = sLanguage
End Function

Function ReadConfiguration(sNode As String, sProperty As String) As String
	Dim oProvider As Object, oAccess As Object
	Dim aArgs(0) As New com.sun.star.beans.PropertyValue
	On Error GoTo Failed
	oProvider = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
	aArgs(0).Name = "nodepath"
	aArgs(0).Value = sNode
	oAccess = oProvider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", aArgs())
	ReadConfiguration = oAccess.getByName(sProperty)
	Exit Function
Failed:
	ReadConfiguration = ""
End Function

' Installation folder of the extension ("" if not found)
Function ExtensionFolder() As String
	Dim s As String
	On Error GoTo Failed
	s = GetDefaultContext().getByName("/singletons/com.sun.star.deployment.PackageInformationProvider") _
		.getPackageLocation(ID_EXTENSION)
	If s <> "" Then s = s & "/"
	ExtensionFolder = s
	Exit Function
Failed:
	ExtensionFolder = ""
End Function

' Reads a key=value file (UTF-8, # comment, \n line break)
' and adds its entries to oTable (EnumerableMap string -> string)
Sub ReadProperties(sUrl As String, oTable As Object)
	Dim aLines As Variant, sLine As String, i As Long, p As Long
	On Error GoTo Failed
	If Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(sUrl) Then Exit Sub
	aLines = Split(ReadText(sUrl), Chr(10))
	For i = 0 To UBound(aLines)
		sLine = LTrim(Replace(aLines(i), Chr(13), ""))
		p = InStr(sLine, "=")
		If p > 1 And Left(sLine, 1) <> "#" Then
			oTable.put(Trim(Left(sLine, p - 1)), Replace(LTrim(Mid(sLine, p + 1)), "\n", Chr(10)))
		End If
	Next i
	Exit Sub
Failed:
	WriteLog("ERROR reading " & ConvertFromURL(sUrl) & ": " & Error$)
End Sub

Sub WriteText(sUrl As String, sText As String)
	Dim oSFA As Object, oStream As Object
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sUrl) Then oSFA.kill(sUrl)
	oStream = CreateUnoService("com.sun.star.io.TextOutputStream")
	oStream.setOutputStream(oSFA.openFileWrite(sUrl))
	oStream.setEncoding("UTF-8")
	oStream.writeString(sText)
	oStream.closeOutput()
End Sub

' ---------------------------------------------------------------------------
' Minimal JSON reader. Object -> Array("{", keys(), values());
' array -> Array("[", elements()). Empty if the text is invalid.
' ---------------------------------------------------------------------------
Function JsonRead(s As String) As Variant
	sJson = s
	nJson = 1
	bJsonOk = True
	JsonRead = JsonValue()
	JsonSpaces()
	If nJson <= Len(sJson) Then bJsonOk = False
	If Not bJsonOk Then
		WriteLog("Invalid JSON near character " & nJson)
		JsonRead = Empty
	End If
	sJson = ""
End Function

Function JsonField(o As Variant, sKey As String) As Variant
	Dim aKeys As Variant, i As Long
	JsonField = Empty
	If Not IsArray(o) Then Exit Function
	If o(0) <> "{" Then Exit Function
	aKeys = o(1)
	For i = 0 To UBound(aKeys)
		If aKeys(i) = sKey Then
			JsonField = o(2)(i)
			Exit Function
		End If
	Next i
End Function

Function JsonElements(o As Variant) As Variant
	JsonElements = Array()
	If Not IsArray(o) Then Exit Function
	If o(0) = "[" Then JsonElements = o(1)
End Function

Function JsonText(o As Variant, sKey As String) As String
	Dim v As Variant
	v = JsonField(o, sKey)
	If VarType(v) = 8 Then JsonText = v Else JsonText = ""
End Function

' Translated field if it exists (e.g. "title_en"), otherwise the base field
Function JsonTranslatedText(o As Variant, sKey As String) As String
	Dim s As String
	s = JsonText(o, sKey & "_" & UiLanguage())
	If s = "" Then s = JsonText(o, sKey)
	JsonTranslatedText = s
End Function

Function JsonNumber(o As Variant, sKey As String) As Double
	Dim v As Variant
	v = JsonField(o, sKey)
	If VarType(v) = 5 Then JsonNumber = v Else JsonNumber = 0
End Function

Sub JsonSpaces()
	Dim c As String
	Do While nJson <= Len(sJson)
		c = Mid(sJson, nJson, 1)
		If c <> " " And c <> Chr(9) And c <> Chr(10) And c <> Chr(13) Then Exit Do
		nJson = nJson + 1
	Loop
End Sub

Function JsonValue() As Variant
	Dim c As String
	JsonValue = Empty
	JsonSpaces()
	If nJson > Len(sJson) Then
		bJsonOk = False
		Exit Function
	End If
	c = Mid(sJson, nJson, 1)
	If c = "{" Then
		JsonValue = JsonObject()
	ElseIf c = "[" Then
		JsonValue = JsonArray()
	ElseIf c = """" Then
		JsonValue = JsonString()
	ElseIf c = "t" Then
		JsonValue = JsonWord("true", True)
	ElseIf c = "f" Then
		JsonValue = JsonWord("false", False)
	ElseIf c = "n" Then
		JsonValue = JsonWord("null", Null)
	Else
		JsonValue = JsonNum()
	End If
End Function

Function JsonObject() As Variant
	Dim aKeys() As Variant, aValues() As Variant, n As Long, c As String
	nJson = nJson + 1
	JsonSpaces()
	If Mid(sJson, nJson, 1) = "}" Then
		nJson = nJson + 1
		JsonObject = Array("{", aKeys, aValues)
		Exit Function
	End If
	Do
		JsonSpaces()
		If Mid(sJson, nJson, 1) <> """" Then
			bJsonOk = False
			Exit Function
		End If
		ReDim Preserve aKeys(n)
		ReDim Preserve aValues(n)
		aKeys(n) = JsonString()
		JsonSpaces()
		If Mid(sJson, nJson, 1) <> ":" Then
			bJsonOk = False
			Exit Function
		End If
		nJson = nJson + 1
		aValues(n) = JsonValue()
		If Not bJsonOk Then Exit Function
		n = n + 1
		JsonSpaces()
		c = Mid(sJson, nJson, 1)
		nJson = nJson + 1
		If c = "}" Then Exit Do
		If c <> "," Then
			bJsonOk = False
			Exit Function
		End If
	Loop
	JsonObject = Array("{", aKeys, aValues)
End Function

Function JsonArray() As Variant
	Dim aElements() As Variant, n As Long, c As String
	nJson = nJson + 1
	JsonSpaces()
	If Mid(sJson, nJson, 1) = "]" Then
		nJson = nJson + 1
		JsonArray = Array("[", aElements)
		Exit Function
	End If
	Do
		ReDim Preserve aElements(n)
		aElements(n) = JsonValue()
		If Not bJsonOk Then Exit Function
		n = n + 1
		JsonSpaces()
		c = Mid(sJson, nJson, 1)
		nJson = nJson + 1
		If c = "]" Then Exit Do
		If c <> "," Then
			bJsonOk = False
			Exit Function
		End If
	Loop
	JsonArray = Array("[", aElements)
End Function

Function JsonString() As String
	Dim s As String, c As String, k As Integer, nCode As Long
	nJson = nJson + 1
	Do While nJson <= Len(sJson)
		c = Mid(sJson, nJson, 1)
		If c = """" Then
			nJson = nJson + 1
			JsonString = s
			Exit Function
		ElseIf c = "\" Then
			nJson = nJson + 1
			c = Mid(sJson, nJson, 1)
			If c = "n" Then
				s = s & Chr(10)
			ElseIf c = "t" Then
				s = s & Chr(9)
			ElseIf c = "r" Then
				s = s & Chr(13)
			ElseIf c = "b" Or c = "f" Then
				' control characters ignored
			ElseIf c = "u" Then
				nCode = 0
				For k = 1 To 4
					nCode = nCode * 16 + InStr(1, "0123456789abcdef", LCase(Mid(sJson, nJson + k, 1)), 0) - 1
				Next k
				If nCode >= 0 Then s = s & Chr(nCode)
				nJson = nJson + 4
			Else
				s = s & c                     ' \" \\ \/
			End If
		Else
			s = s & c
		End If
		nJson = nJson + 1
	Loop
	bJsonOk = False
	JsonString = s
End Function

Function JsonWord(sWord As String, v As Variant) As Variant
	If Mid(sJson, nJson, Len(sWord)) = sWord Then
		nJson = nJson + Len(sWord)
		JsonWord = v
	Else
		bJsonOk = False
		JsonWord = Empty
	End If
End Function

Function JsonNum() As Double
	Dim nStart As Long
	nStart = nJson
	Do While nJson <= Len(sJson)
		If InStr(1, "+-0123456789.eE", Mid(sJson, nJson, 1), 0) = 0 Then Exit Do
		nJson = nJson + 1
	Loop
	If nJson = nStart Then
		bJsonOk = False
		JsonNum = 0
	Else
		JsonNum = Val(Mid(sJson, nStart, nJson - nStart))
	End If
End Function

' ---------------------------------------------------------------------------
' Messages and log
' ---------------------------------------------------------------------------
Sub Warn(sMessage As String)
	On Error Resume Next
	WriteLog("Message: " & sMessage)
	If sLogPath <> "" Then
		MsgBox sMessage & Chr(10) & Chr(10) & Tr("msg.details", sLogPath), 48, Tr("app.title")
	Else
		MsgBox sMessage, 48, Tr("app.title")
	End If
End Sub

Sub Inform(sMessage As String)
	On Error Resume Next
	WriteLog("Information: " & sMessage)
	MsgBox sMessage, 64, Tr("app.title")
End Sub

Sub OpenLog()
	Dim oSFA As Object, oStream As Object, sUrl As String
	On Error GoTo Failed
	oLog = Nothing
	sLogPath = ""
	sUrl = CreateUnoService("com.sun.star.util.PathSubstitution") _
		.substituteVariables("$(user)/" & LOG_NAME, True)
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sUrl) Then oSFA.kill(sUrl)
	oStream = oSFA.openFileWrite(sUrl)
	oLog = CreateUnoService("com.sun.star.io.TextOutputStream")
	oLog.setOutputStream(oStream)
	oLog.setEncoding("UTF-8")
	sLogPath = ConvertFromURL(sUrl)
	Exit Sub
Failed:
	oLog = Nothing        ' no log: the macro works anyway
End Sub

' Adds a line to the log outside a run of the macro
' (listeners: context menu…)
Sub AppendLog(sText As String)
	Dim oSFA As Object, oStream As Object, oText As Object, sUrl As String
	On Error GoTo Failed
	If Not IsNull(oLog) Then
		WriteLog(sText)
		Exit Sub
	End If
	sUrl = CreateUnoService("com.sun.star.util.PathSubstitution") _
		.substituteVariables("$(user)/" & LOG_NAME, True)
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	oStream = oSFA.openFileReadWrite(sUrl)
	oStream.seek(oStream.getLength())
	oText = CreateUnoService("com.sun.star.io.TextOutputStream")
	oText.setOutputStream(oStream.getOutputStream())
	oText.setEncoding("UTF-8")
	oText.writeString(Format(Now, "yyyy-mm-dd hh:mm:ss") & "  " & sText & Chr(10))
	oText.closeOutput()
Failed:
End Sub

Sub WriteLog(sText As String)
	On Error Resume Next
	If IsNull(oLog) Then Exit Sub
	oLog.writeString(Format(Now, "yyyy-mm-dd hh:mm:ss") & "  " & sText & Chr(10))
End Sub

Sub CloseLog()
	On Error Resume Next
	If Not IsNull(oLog) Then oLog.closeOutput()
	oLog = Nothing
End Sub
