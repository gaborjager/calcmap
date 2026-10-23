
' ===========================================================================
' CalcMap test macros, added to a copy of the code by run.sh
' (never included in CalcMap.oxt). Results: "TEST" lines of the log.
' ===========================================================================

Function DocCalcT() As Object
	Dim e As Object, o As Object
	e = StarDesktop.getComponents().createEnumeration()
	Do While e.hasMoreElements()
		o = e.nextElement()
		If o.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then DocCalcT = o
	Loop
End Function

' Inserts the departments map into the 1st sheet, like CalcMap();
' returns the link name
Function InsertMapT(aOpt As Variant, sPalette As String) As String
	Dim oCalc As Object, oDoc As Object, oSel As Object, oRange As Object, aPals As Variant, aPal As Variant
	Dim oShp As Object, bImage As Boolean, sUrl As String, vT As Variant
	Dim aProps(0) As New com.sun.star.beans.PropertyValue
	oCalc = DocCalcT()
	aPals = LoadPalettes()
	aPal = aPals(PaletteIndex(aPals, sPalette))
	oSel = SelectedRange(oCalc.getSheets().getByIndex(0).getCellRangeByName("A1:C1048576"))
	oRange = DataRange(oSel)
	aProps(0).Name = "Hidden" : aProps(0).Value = True
	sUrl = ConvertToURL("/home/gabor/github/france-odg/output/departements-france.odg")
	oDoc = StarDesktop.loadComponentFromURL(sUrl, "_blank", 0, aProps())
	vT = TextContext(oCalc, oRange, oRange.getDataArray(), aOpt, Nothing)
	ColorDocument(oDoc, oRange.getDataArray(), aPal, vT)
	oShp = InsertIntoCalc(oDoc, oCalc, oRange, FileNameOf(sUrl), bImage)
	oDoc.close(True)
	InsertMapT = CreateLink(oCalc, oSel, oShp, aPal(P_ID), sUrl, bImage, aOpt)
	If aOpt(O_LEGEND) Or aOpt(O_VALUES) Or aOpt(O_LABELS) Then RecalcLink(oCalc, InsertMapT)
End Function

' ---------------------------------------------------------------- texts
Function TextsT(oG As Object, sCode As String) As String
	Dim i As Long, nZ As Long, nT As Long, nL As Long, iFirstText As Long, iLastZone As Long
	Dim s As String, o As Object
	iFirstText = -1
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		If Left(o.Name, Len(TEXT_PREFIX)) = TEXT_PREFIX Then
			nT = nT + 1
			If iFirstText < 0 Then iFirstText = i
			If o.Name = TEXT_PREFIX & " " & sCode Then s = s & "text '" & Replace(o.getString(), Chr(10), "/") _
				& "' color " & Hex(o.CharColor) & " center " & (o.Position.X + o.Size.Width / 2) & "," _
				& (o.Position.Y + o.Size.Height / 2)
		ElseIf IsLegend(o.Name) Then
			nL = nL + 1
		Else
			nZ = nZ + 1
			iLastZone = i
			If o.Name = sCode Then s = s & " ; zone center " & (o.Position.X + o.Size.Width / 2) & "," _
				& (o.Position.Y + o.Size.Height / 2) & " text in the zone='" & o.getString() & "'"
		End If
	Next i
	TextsT = nZ & " zones, " & nT & " texts, " & nL & " legend shapes ; texts after the zones: " _
		& (nT = 0 Or iFirstText > iLastZone) & " | " & sCode & ": " & s
End Function

Sub TestTexts()
	Dim oCalc As Object, oSheet As Object, oG As Object, aOpt As Variant, sLink As String, i As Long, t As Long
	Dim aFields As Variant, oDoc As Object, oRange As Object, vT As Variant, aPals As Variant
	Dim aProps(0) As New com.sun.star.beans.PropertyValue
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	oCalc = DocCalcT()
	oSheet = oCalc.getSheets().getByIndex(0)
	aOpt = DefaultOptions()
	aOpt(O_VALUES) = True : aOpt(O_LABELS) = True : aOpt(O_LEGEND) = True
	aOpt(O_SIZE_VALUES) = 14 : aOpt(O_SIZE_LABELS) = 7
	t = GetSystemTicks()
	sLink = InsertMapT(aOpt, "jaune-rouge")
	oG = oSheet.getDrawPage().getByIndex(0)
	WriteLog("TEST insertion (" & (GetSystemTicks() - t) & " ms): " & TextsT(oG, "88"))
	' Change of a value: texts up to date, no duplicates
	t = GetSystemTicks()
	For i = 1 To 200
		If oSheet.getCellByPosition(0, i).getString() = "88" Then oSheet.getCellByPosition(1, i).setValue(4321)
	Next i
	WriteLog("TEST change (" & (GetSystemTicks() - t) & " ms): " & TextsT(oG, "88"))
	' Map from a previous version: text in the zone, removed
	For i = 0 To oG.getCount() - 1
		If oG.getByIndex(i).Name = "88" Then oG.getByIndex(i).setString("old")
	Next i
	RecalcLink(oCalc, sLink, True)
	WriteLog("TEST old text in the zone: " & TextsT(oG, "88"))
	' Values and labels disabled: texts removed
	aFields = Split(oCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sLink), "|")
	aOpt(O_VALUES) = False : aOpt(O_LABELS) = False
	oCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sLink, _
		LinkFields(aFields(1), aFields(2), aFields(3), False, aOpt))
	RecalcLink(oCalc, sLink)
	WriteLog("TEST options disabled: " & TextsT(oG, "88"))
	aOpt(O_VALUES) = True : aOpt(O_LABELS) = True
	oCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sLink, _
		LinkFields(aFields(1), aFields(2), aFields(3), False, aOpt))
	RecalcLink(oCalc, sLink)
	' Opening in Draw
	aPals = LoadPalettes()
	oRange = DataRange(SelectedRange(oSheet.getCellRangeByName("A1:C1048576")))
	aProps(0).Name = "Hidden" : aProps(0).Value = True
	oDoc = StarDesktop.loadComponentFromURL(ConvertToURL("/home/gabor/github/france-odg/output/departements-france.odg"), _
		"_blank", 0, aProps())
	vT = TextContext(oCalc, oRange, oRange.getDataArray(), aOpt, oDoc)
	ColorDocument(oDoc, oRange.getDataArray(), aPals(PaletteIndex(aPals, "jaune-rouge")), vT)
	WriteLog("TEST Draw: " & TextsT(oDoc.getDrawPages().getByIndex(0), "88"))
	oDoc.close(True)
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- classes editor
Function StateT() As String
	Dim k As Integer, s As String, aN As Variant
	aN = CountClasses()
	For k = 0 To nEdClasses - 1
		s = s & "[" & EdNumber(aEdBounds(k)) & ".." & EdNumber(aEdBounds(k + 1)) & " " & Hex(aEdColors(k)) _
			& " n=" & aN(k) & "]"
	Next k
	StateT = nEdClasses & " classes (field " & oEditor.getControl("numClasses").getText() & ") " & s _
		& " | chosen bound=" & iEdBound & " field='" & oEditor.getControl("txtBound").getText() & "'" _
		& " | error='" & oEditor.getControl("lblError").getModel().Label & "'"
End Function

Sub MouseT(sWhat As String, sName As String, nX As Long)
	Dim oEv As Object
	oEv = CreateUnoStruct("com.sun.star.awt.MouseEvent")
	oEv.Source = oEditor.getControl(sName)
	oEv.X = nX
	oEv.Buttons = com.sun.star.awt.MouseButton.LEFT
	Select Case sWhat
		Case "press" : Bar_mousePressed(oEv)
		Case "drag" : BarMove_mouseDragged(oEv)
		Case "release" : Bar_mouseReleased(oEv)
	End Select
End Sub

Sub OpenEditorT(sPalette As String)
	Dim aPals As Variant, oRange As Object, s As String
	LoadTexts()
	aPals = LoadPalettes()
	oRange = DataRange(SelectedRange(DocCalcT().getSheets().getByIndex(0).getCellRangeByName("A1:C1048576")))
	s = BuildEditor("", aPals(PaletteIndex(aPals, sPalette)), oRange.getDataArray(), 0, 1, 2, Nothing)
End Sub

Sub TestFullEditor()
	Dim oBase As Object, oEv As Object, oNum As Object, k As Integer
	OpenLog()
	On Error GoTo ErrT
	OpenEditorT("bleus")
	oBase = oEditor.getControl("barBase").getPosSize()
	oNum = oEditor.getControl("numClasses")
	WriteLog("TEST start: " & StateT())
	oNum.up() : oNum.up()
	WriteLog("TEST 2 clicks up: " & StateT())
	MouseT("press", "pgn3", 1)
	MouseT("drag", "barBase", oBase.Width / 2)
	MouseT("release", "barBase", oBase.Width / 2)
	WriteLog("TEST bound 3 dragged to the middle: " & StateT())
	MouseT("press", "pgn3", 1)
	MouseT("drag", "barBase", oBase.Width - 2)
	MouseT("release", "barBase", oBase.Width - 2)
	WriteLog("TEST bound 3 dragged too far: " & StateT())
	oEditor.getControl("txtBound").setText("450")
	WriteLog("TEST bound 3 typed 450: " & StateT())
	oEditor.getControl("txtBound").setText("5000")
	WriteLog("TEST bound 3 typed 5000: " & StateT())
	aEdColors(2) = RGB(255, 0, 0)
	oEv = CreateUnoStruct("com.sun.star.awt.ActionEvent")
	oEv.Source = oEditor.getControl("btnLinear")
	EditorButton_actionPerformed(oEv)
	WriteLog("TEST linear (7 classes and red kept): " & StateT())
	oNum.setText("1")
	WriteLog("TEST typed '1' (ignored): " & StateT())
	oNum.setText("12")
	WriteLog("TEST typed '12': " & StateT())
	For k = 1 To 4 : oNum.down() : Next k
	WriteLog("TEST 4 clicks down: " & StateT())
	For k = 1 To 30 : oNum.up() : Next k
	WriteLog("TEST 30 clicks up (max 20): " & StateT())
	MouseT("press", "seg1", 3)
	MouseT("release", "seg1", 3)
	WriteLog("TEST click on a color: " & StateT())
	oEditor.dispose()
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- readability
' Departments map with each readability setting: number of shapes,
' recoloring time, image readability-<setting>.png
Sub TestMapReadability()
	Dim oCalc As Object, oSheet As Object, aOpt As Variant, sLink As String, sMode As Variant, t As Long
	Dim oG As Object, i As Long, aFields As Variant
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	oCalc = DocCalcT()
	oSheet = oCalc.getSheets().getByIndex(0)
	aOpt = DefaultOptions()
	WriteLog("TEST default outline: values " & aOpt(O_OUTLINE_VALUES) & ", labels " & aOpt(O_OUTLINE_LABELS))
	aOpt(O_VALUES) = True : aOpt(O_LABELS) = True : aOpt(O_SIZE_LABELS) = 7
	sLink = InsertMapT(aOpt, "jaune-rouge-10")
	oG = oSheet.getDrawPage().getByIndex(0)
	aFields = Split(oCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sLink), "|")
	For Each sMode In Array("none", "outline", "values only")
		aOpt(O_OUTLINE_VALUES) = IIf(sMode = "none", "none", "outline")
		aOpt(O_OUTLINE_LABELS) = IIf(sMode = "outline", "outline", "none")
		oCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sLink, _
			LinkFields(aFields(1), aFields(2), aFields(3), False, aOpt))
		t = GetSystemTicks()
		RecalcLink(oCalc, sLink)
		WriteLog("TEST " & sMode & ": " & oG.getCount() & " shapes in the map, recoloring " & (GetSystemTicks() - t) & " ms ; " & CopyT(oG))
		ExportShapeT(oG, "readability-" & sMode & ".png")
	Next sMode
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

Sub ExportShapeT(oShp As Object, sName As String)
	Dim oFilter As Object
	Dim aArgs(2) As New com.sun.star.beans.PropertyValue, aDataArray(0) As New com.sun.star.beans.PropertyValue
	aDataArray(0).Name = "PixelWidth" : aDataArray(0).Value = 2000
	aArgs(0).Name = "URL" : aArgs(0).Value = ConvertToURL(Environ("CALCMAP_TESTS") & "/" & sName)
	aArgs(1).Name = "MediaType" : aArgs(1).Value = "image/png"
	aArgs(2).Name = "FilterData" : aArgs(2).Value = aDataArray()
	oFilter = CreateUnoService("com.sun.star.drawing.GraphicExportFilter")
	oFilter.setSourceDocument(oShp)
	oFilter.filter(aArgs())
End Sub

' ---------------------------------------------------------------- main dialog
' The dialog is not shown without a display: it is built, then its
' settings are read back; they must come back unchanged
Sub TestDialog()
	Dim aOpt As Variant, aPals As Variant, iC As Integer, iP As Integer, bI As Boolean, n As Integer, sBefore As String
	Dim oRange As Object, sMode As Variant
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	aPals = LoadPalettes()
	oRange = DataRange(SelectedRange(DocCalcT().getSheets().getByIndex(0).getCellRangeByName("A1:C1048576")))
	aDialogData = oRange.getDataArray()
	For Each sMode In Array("none", "outline")
		aOpt = DefaultOptions()
		aOpt(O_VALUES) = True : aOpt(O_STYLE_LABELS) = "BI" : aOpt(O_SIZE_LABELS) = 13
		aOpt(O_OUTLINE_VALUES) = sMode : aOpt(O_OUTLINE_LABELS) = IIf(sMode = "none", "outline", "none")
		sBefore = OptionsText(aOpt)
		bI = True
		n = MapDialog(Array(NewMap("departements", "Dep", "d", "", "d.odg", 1, "", True)), aPals, "state", _
			"", "bleus", iC, iP, bI, aOpt, ColumnNames(oRange), Nothing, False)
		WriteLog("TEST dialog " & sMode & ": read back unchanged=" & (OptionsText(aOpt) = sBefore) & " ; " & OptionsText(aOpt))
	Next sMode
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- partial update
' Map state: color, alternative text and texts (sorted) of each zone
Function MapStateT(oG As Object) As String
	Dim i As Long, o As Object, aT() As String, n As Long
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		ReDim Preserve aT(n) As String
		If IsLegend(o.Name) Then
			aT(n) = "L " & o.getShapeType() & " " & o.Position.X & "," & o.Position.Y & IIf(o.getShapeType() = "com.sun.star.drawing.TextShape", " '" & o.getString() & "'", " " & Hex(o.FillColor))
		ElseIf Left(o.Name, Len(TEXT_PREFIX)) = TEXT_PREFIX Then
			aT(n) = "T " & o.Name & " '" & Replace(o.getString(), Chr(10), "/") & "' " & Hex(o.CharColor) & " " & o.Position.X & "," & o.Position.Y
		Else
			aT(n) = "Z " & o.Name & " " & Hex(o.FillColor) & " [" & o.Description & "]"
		End If
		n = n + 1
	Next i
	SortTexts(aT)
	MapStateT = Join(aT, Chr(10))
End Function

' Changes cells (address=value; empty: cell cleared) then compares
' the update (partial or not) with a full recalculation
Sub PartialCaseT(sScenario As String, aEdits As Variant, sLink As String, oG As Object)
	Dim oCalc As Object, oSheet As Object, k As Integer, a As Variant, t As Long, sAfter As String, sFull As String
	Dim nZones As Long
	oCalc = DocCalcT()
	oSheet = oCalc.getSheets().getByIndex(0)
	t = GetSystemTicks()
	For k = 0 To UBound(aEdits)
		a = Split(aEdits(k), "=")
		If a(1) = "" Then
			oSheet.getCellRangeByName(a(0)).setString("")
		ElseIf IsNumeric(a(1)) Then
			oSheet.getCellRangeByName(a(0)).setValue(CDbl(a(1)))
		Else
			oSheet.getCellRangeByName(a(0)).setString(a(1))
		End If
	Next k
	t = GetSystemTicks() - t
	nZones = nLastUpdate
	sAfter = MapStateT(oG)
	RecalcLink(oCalc, sLink, True)
	sFull = MapStateT(oG)
	WriteLog("TEST " & sScenario & ": " & IIf(nZones < 0, "whole map", nZones & " zone(s)") & " in " & t & " ms ; same as full recalculation: " & (sAfter = sFull))
	If sAfter <> sFull Then WriteLog("TEST   difference : " & DifferenceT(sAfter, sFull))
End Sub

Function DifferenceT(a As String, b As String) As String
	Dim aa As Variant, bb As Variant, i As Long
	aa = Split(a, Chr(10)) : bb = Split(b, Chr(10))
	For i = 0 To UBound(aa)
		If i > UBound(bb) Then DifferenceT = "extra: " & aa(i) : Exit Function
		If aa(i) <> bb(i) Then DifferenceT = aa(i) & "  <>  " & bb(i) : Exit Function
	Next i
	DifferenceT = "number of shapes " & (UBound(aa) + 1) & " / " & (UBound(bb) + 1)
End Function

Sub TestPartial()
	Dim oCalc As Object, oSheet As Object, aOpt As Variant, sLink As String, oG As Object, i As Long, aEdits() As String
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	oCalc = DocCalcT()
	oSheet = oCalc.getSheets().getByIndex(0)
	aOpt = DefaultOptions()
	aOpt(O_VALUES) = True : aOpt(O_LABELS) = True : aOpt(O_LEGEND) = True
	sLink = InsertMapT(aOpt, "jaune-rouge")
	oG = oSheet.getDrawPage().getByIndex(0)
	' Rows: A2 = 01 (Ain, 389) ...; 88 in row 90; min 10 (Lozère, 48), max 988 (Gard, 30)
	PartialCaseT("value changed within the bounds (Ain 389 -> 700)", Array("B2=700"), sLink, oG)
	PartialCaseT("same value typed again", Array("B2=700"), sLink, oG)
	PartialCaseT("label changed (Ain -> AIN)", Array("C2=AIN"), sLink, oG)
	PartialCaseT("code changed (01 -> 02, duplicate)", Array("A2=02"), sLink, oG)
	PartialCaseT("code restored", Array("A2=01"), sLink, oG)
	PartialCaseT("new maximum (Ain -> 2000)", Array("B2=2000"), sLink, oG)
	PartialCaseT("value cleared (Ain)", Array("B2="), sLink, oG)
	PartialCaseT("value restored (Ain 389)", Array("B2=389"), sLink, oG)
	ReDim aEdits(24) As String
	For i = 0 To 24
		aEdits(i) = "B" & (i + 3) & "=" & (500 + i)
	Next i
	PartialCaseT("25 values changed one by one", aEdits, sLink, oG)
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' Paragraphs of an outline copy of the Vosges text: text and transparency
Function CopyT(oG As Object) As String
	Dim i As Long, o As Object, oEnum As Object, p As Object, s As String
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		If o.Name = TEXT_PREFIX & " 88" Then
			oEnum = o.getText().createEnumeration()
			s = ""
			Do While oEnum.hasMoreElements()
				p = oEnum.nextElement()
				s = s & "[" & p.getString() & "]"
			Loop
			CopyT = "1st text shape of Vosges: " & s & " color " & Hex(o.CharColor)
			Exit Function
		End If
	Next i
End Function

' ---------------------------------------------------------------- previews
' Is an SVG image shown in a dialog image control?
Sub TestSvgPreview()
	Dim oModel As Object, oDlg As Object, oImg As Object, oGraphic As Object
	Dim aSource(0) As New com.sun.star.beans.PropertyValue
	OpenLog()
	On Error GoTo ErrT
	aSource(0).Name = "URL" : aSource(0).Value = ExtensionFolder() & "icons/nopreview.svg"
	oGraphic = CreateUnoService("com.sun.star.graphic.GraphicProvider").queryGraphic(aSource())
	WriteLog("TEST svg loaded: " & Not IsNull(oGraphic) & " ; type " & oGraphic.GraphicType & " ; size " & oGraphic.Size100thMM.Width & "x" & oGraphic.Size100thMM.Height)
	oModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
	oModel.Width = 200 : oModel.Height = 150
	AddControl(oModel, "ImageControl", "img", 10, 10, 100, 80, Array("ScaleImage", "ScaleMode", "ImageURL"), _
		Array(True, com.sun.star.awt.ImageScaleMode.ISOTROPIC, aSource(0).Value))
	oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
	oDlg.setModel(oModel)
	oDlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Nothing)
	WriteLog("TEST image control: graphic present=" & Not IsNull(oDlg.getControl("img").getModel().Graphic))
	oDlg.dispose()
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- catalog
Sub SettingT(sName As String, v As Variant)
	Dim oAccess As Object, aArgs(0) As New com.sun.star.beans.PropertyValue
	aArgs(0).Name = "nodepath" : aArgs(0).Value = SETTINGS_NODE
	oAccess = CreateUnoService("com.sun.star.configuration.ConfigurationProvider") _
		.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", aArgs())
	oAccess.setPropertyValue(sName, v)
	oAccess.commitChanges()
End Sub

Function ListT(aMaps As Variant) As String
	Dim i As Long, s As String
	For i = 0 To UBound(aMaps)
		s = s & aMaps(i)(C_ID) & "[" & aMaps(i)(C_SOURCE) & ", " & aMaps(i)(C_OBJECTS) & " zones" _
			& IIf(aMaps(i)(C_PREVIEW) <> "", ", preview", "") & IIf(aMaps(i)(C_CODES) <> "", ", codes " & aMaps(i)(C_CODES), "") & "] "
	Next i
	ListT = (UBound(aMaps) + 1) & " map(s): " & s
End Function

Sub TestCatalog()
	Dim sTests As String, sFixtures As String, sFrance As String, s As String, aMaps As Variant, b As Boolean, i As Long
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	WriteLog("TEST default settings: catalog=" & CatalogUrl() & " ; max zones=" & MaxZones())
	sTests = ConvertToURL(Environ("CALCMAP_TESTS"))
	sFixtures = ConvertToURL(Environ("CALCMAP_FIXTURES"))
	sFrance = ConvertToURL("/home/gabor/github/france-odg/output")
	' france-odg source: the project's catalog, previews as local files
	s = Replace(ReadText(sFrance & "/maps.json"), "https://raw.githubusercontent.com/gaborjager/france-odg/main/output", sFrance)
	WriteText(sTests & "/source-france.json", s)
	WriteText(sTests & "/source-company.json", Replace(Replace(ReadText(sFixtures & "/source-company.json"), _
		"@TESTS@", sTests), "@FRANCEODG@", sFrance))
	WriteText(sTests & "/aggregate.json", Replace(ReadText(sFixtures & "/aggregate.json"), "@TESTS@", sTests))
	SettingT("Catalog", sTests & "/aggregate.json")
	SettingT("MaxZones", 200)
	aMaps = LoadCatalog(MapsFolder(), b)
	WriteLog("TEST max 200: online=" & b & " ; " & ListT(aMaps))
	SettingT("MaxZones", 3000)
	aMaps = LoadCatalog(MapsFolder(), b)
	WriteLog("TEST max 3000: " & ListT(aMaps))
	For i = 0 To UBound(aMaps)
		If aMaps(i)(C_ID) = "departements" Or aMaps(i)(C_ID) = "secteurs" Then
			s = LocalPreview(aMaps(i))
			WriteLog("TEST preview " & aMaps(i)(C_ID) & ": '" & s & "' readable=" & Not IsNull(GraphicOf(s)))
		End If
	Next i
	WriteLog("TEST default image readable: " & Not IsNull(GraphicOf(ExtensionFolder() & "icons/nopreview.svg")))
	' Catalog unreachable: cached copy
	SettingT("Catalog", sTests & "/absent.json")
	aMaps = LoadCatalog(MapsFolder(), b)
	WriteLog("TEST catalog unreachable: online=" & b & " ; " & ListT(aMaps))
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- settings
' Settings read, then changed as in the expert configuration
Sub TestSettings()
	OpenLog()
	WriteLog("TEST settings read: catalog=" & CatalogUrl() & " ; max zones=" & MaxZones())
	On Error GoTo Refused
	SettingT("MaxZones", 7)
	WriteLog("TEST change by the user accepted: max zones=" & MaxZones())
	CloseLog()
	Exit Sub
Refused:
	WriteLog("TEST change by the user refused (" & Error$ & "): max zones=" & MaxZones())
	CloseLog()
End Sub

' ---------------------------------------------------------------- icon on a map
' Map selected (or one of its zones): the icon opens its configuration
' (the dialog is not shown without a display) and does not insert a new map
Sub TestIconOnMap()
	Dim oCalc As Object, oSheet As Object, aOpt As Variant, sLink As String, oG As Object, n As Long
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	oCalc = DocCalcT()
	oSheet = oCalc.getSheets().getByIndex(0)
	aOpt = DefaultOptions()
	sLink = InsertMapT(aOpt, "bleus")
	oG = oSheet.getDrawPage().getByIndex(0)
	CloseLog()
	oCalc.getCurrentController().select(oG)
	CalcMap()
	n = oSheet.getDrawPage().getCount()
	oCalc.getCurrentController().select(oG.getByIndex(10))
	CalcMap()
	AppendLog("TEST after the icon on the map then on a zone: " & oSheet.getDrawPage().getCount() & " shape(s) in the sheet (before: " & n & ")")
	Exit Sub
ErrT:
	AppendLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub


' ---------------------------------------------------------------- translations
' Each language: file loaded, all keys translated, example read as
' UTF-8; dialog built in a few languages
Sub TestTranslations()
	Dim aCodes As Variant, sCode As Variant, aKeys As Variant, sKey As Variant, nMissing As Integer, s As String
	Dim oRange As Object, aOpt As Variant, aPals As Variant, iC As Integer, iP As Integer, bI As Boolean, n As Integer
	OpenLog()
	On Error GoTo ErrT
	aCodes = Split("en fr ar bg cs da de el es et fi ga hi hr hu id it ja ko lt lv mt nb nl pl pt ro ru sk sl sv tr uk zh", " ")
	aKeys = Array("dlg.grp_map", "dlg.show_values", "ed.help", "msg.inserted_as_image", "dlg.zones", "menu.configure")
	For Each sCode In aCodes
		sLanguage = sCode
		LoadTexts()
		nMissing = 0
		For Each sKey In aKeys
			If Tr(sKey) = sKey Then nMissing = nMissing + 1
		Next sKey
		s = s & sCode & "=" & Tr("dlg.grp_map") & IIf(nMissing > 0, "(!" & nMissing & ")", "") & " "
	Next sCode
	WriteLog("TEST languages: " & s)
	sLanguage = "fr" : LoadTexts()
	WriteLog("TEST French differs from English: " & Tr("dlg.show_values") & " / %3 : " & Tr("ed.info", 1, 2, 96))
	aPals = LoadPalettes()
	oRange = DataRange(SelectedRange(DocCalcT().getSheets().getByIndex(0).getCellRangeByName("A1:C1048576")))
	aDialogData = oRange.getDataArray()
	For Each sCode In Array("de", "ar", "ja")
		sLanguage = sCode : LoadTexts()
		aOpt = DefaultOptions()
		bI = True
		n = MapDialog(Array(NewMap("departements", "Dep", "d", "", "d.odg", 1, "", True)), aPals, "", _
			"", "bleus", iC, iP, bI, aOpt, ColumnNames(oRange), Nothing, False)
		WriteLog("TEST dialog in " & sCode & ": built (result " & n & ")")
	Next sCode
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- columns
' Column names: first 3 values, ", ..." only if more follow
Sub TestColumnNames()
	Dim oSheet As Object
	OpenLog()
	On Error GoTo ErrT
	oSheet = DocCalcT().getSheets().getByIndex(0)
	WriteLog("TEST 96 rows: " & Join(ColumnNames(DataRange(SelectedRange(oSheet.getCellRangeByName("A1:C1048576")))), " | "))
	WriteLog("TEST 3 rows: " & Join(ColumnNames(oSheet.getCellRangeByName("A1:C4")), " | "))
	WriteLog("TEST 2 rows: " & Join(ColumnNames(oSheet.getCellRangeByName("A1:C3")), " | "))
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- codes icon, geometry
' Map codes looked up in the code column: tick, then triangle
' when another column is chosen; controls inside the dialog, without
' overlap outside the frames
Sub TestCodesIcon()
	Dim aOpt As Variant, aPals As Variant, iC As Integer, iP As Integer, bI As Boolean, oRange As Object
	Dim oImg As Object, aCtrl As Variant, i As Integer, j As Integer, a As Object, b As Object, n As Integer
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	aPals = LoadPalettes()
	oRange = DataRange(SelectedRange(DocCalcT().getSheets().getByIndex(0).getCellRangeByName("A1:C1048576")))
	aDialogData = oRange.getDataArray()
	aOpt = DefaultOptions()
	bI = True
	bKeepDialog = True
	n = MapDialog(Array(NewMap("departements", "Dep", "d", "", "d.odg", 1, "", True, 101, "", "", "", "01, 02, 2A, 2B")), _
		aPals, "state", "", "bleus", iC, iP, bI, aOpt, ColumnNames(oRange), Nothing, False)
	oImg = oDialog.getControl("imgCodes")
	WriteLog("TEST icon, code column: visible=" & oImg.isVisible() & " help=" & oImg.getModel().HelpText)
	oDialog.getControl("lstCode").selectItemPos(2, True)
	CodeList_itemStateChanged(Nothing)
	WriteLog("TEST icon, name column: help=" & oImg.getModel().HelpText)
	n = 0
	aCtrl = oDialog.getModel().getControlModels()
	For i = 0 To UBound(aCtrl)
		a = aCtrl(i)
		If a.PositionX + a.Width > oDialog.getModel().Width Or a.PositionY + a.Height > oDialog.getModel().Height Then
			WriteLog("TEST geometry outside the dialog: " & a.Name) : n = n + 1
		End If
		For j = i + 1 To UBound(aCtrl)
			b = aCtrl(j)
			' Frames, title (in place of the list) and unused (hidden) shades excluded
			If InStr(a.Name & b.Name, "grp") = 0 And InStr(a.Name & b.Name, "lblMapTitle") = 0 And _
				Not (Left(a.Name, 5) = "shade" And Left(b.Name, 5) = "shade") And _
				a.PositionX < b.PositionX + b.Width And b.PositionX < a.PositionX + a.Width And _
				a.PositionY < b.PositionY + b.Height And b.PositionY < a.PositionY + a.Height Then
				WriteLog("TEST geometry overlap: " & a.Name & " / " & b.Name) : n = n + 1
			End If
		Next j
	Next i
	WriteLog("TEST geometry: " & (UBound(aCtrl) + 1) & " controls, " & n & " problem(s)")
	bKeepDialog = False
	oDialog.dispose()
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- names of older versions
' Map made by a version up to 0.19: French shape names and "aucune" outline;
' recoloring (full, then partial) replaces them without duplicates
Function OldNamesT(oG As Object) As String
	Dim i As Long, nOld As Long, nText As Long, nLegend As Long, sName As String
	For i = 0 To oG.getCount() - 1
		sName = oG.getByIndex(i).Name
		If Left(sName, Len(OLD_TEXT_PREFIX)) = OLD_TEXT_PREFIX Or Left(sName, Len(OLD_LEGEND_PREFIX)) = OLD_LEGEND_PREFIX Then
			nOld = nOld + 1
		ElseIf Left(sName, Len(TEXT_PREFIX)) = TEXT_PREFIX Then
			nText = nText + 1
		ElseIf IsLegend(sName) Then
			nLegend = nLegend + 1
		End If
	Next i
	OldNamesT = nText & " texts, " & nLegend & " legend shapes, " & nOld & " old names"
End Function

Sub TestOldNames()
	Dim oCalc As Object, oSheet As Object, oG As Object, aOpt As Variant, sLink As String, i As Long, o As Object
	OpenLog()
	On Error GoTo ErrT
	LoadTexts()
	oCalc = DocCalcT()
	oSheet = oCalc.getSheets().getByIndex(0)
	aOpt = DefaultOptions()
	aOpt(O_VALUES) = True : aOpt(O_LABELS) = True : aOpt(O_LEGEND) = True
	aOpt(O_OUTLINE_VALUES) = OLD_OUTLINE_NONE : aOpt(O_OUTLINE_LABELS) = OLD_OUTLINE_NONE
	sLink = InsertMapT(aOpt, "bleus")
	oG = oSheet.getDrawPage().getByIndex(0)
	WriteLog("TEST inserted with outline 'aucune': " & OldNamesT(oG))
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		If Left(o.Name, Len(TEXT_PREFIX)) = TEXT_PREFIX Then
			o.Name = OLD_TEXT_PREFIX & Mid(o.Name, Len(TEXT_PREFIX) + 1)
		ElseIf IsLegend(o.Name) Then
			o.Name = OLD_LEGEND_PREFIX
		End If
	Next i
	WriteLog("TEST renamed as in 0.19: " & OldNamesT(oG))
	RecalcLink(oCalc, sLink, True)
	WriteLog("TEST full recoloring: " & OldNamesT(oG))
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		If o.Name = TEXT_PREFIX & " 88" Then o.Name = OLD_TEXT_PREFIX & " 88"
	Next i
	For i = 1 To 200
		If oSheet.getCellByPosition(0, i).getString() = "88" Then oSheet.getCellByPosition(1, i).setValue(500)
	Next i
	WriteLog("TEST partial update of 88 (" & nLastUpdate & " zone(s)): " & OldNamesT(oG))
	CloseLog()
	Exit Sub
ErrT:
	WriteLog("TEST ERROR line " & Erl & ": " & Error$)
	Resume Next
End Sub
