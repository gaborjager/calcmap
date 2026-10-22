
' ===========================================================================
' Macros de test de CalcMap, ajoutées à une copie du code par lancer.sh
' (jamais incluses dans CalcMap.oxt). Résultats : lignes « TEST » du journal.
' ===========================================================================

Function DocCalcT() As Object
	Dim e As Object, o As Object
	e = StarDesktop.getComponents().createEnumeration()
	Do While e.hasMoreElements()
		o = e.nextElement()
		If o.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then DocCalcT = o
	Loop
End Function

' Insère la carte des départements dans la 1re feuille, comme CalcMap() ;
' renvoie le nom du lien
Function InsererCarteT(aOpt As Variant, sPalette As String) As String
	Dim oCalc As Object, oDoc As Object, oSel As Object, oPlage As Object, aPals As Variant, aPal As Variant
	Dim oForme As Object, bImage As Boolean, sUrl As String, vT As Variant
	Dim aProps(0) As New com.sun.star.beans.PropertyValue
	oCalc = DocCalcT()
	aPals = ChargerPalettes()
	aPal = aPals(IndicePalette(aPals, sPalette))
	oSel = PlageSelectionnee(oCalc.getSheets().getByIndex(0).getCellRangeByName("A1:C1048576"))
	oPlage = PlageDeDonnees(oSel)
	aProps(0).Name = "Hidden" : aProps(0).Value = True
	sUrl = ConvertToURL("/home/gabor/github/france-odg/output/departements-france.odg")
	oDoc = StarDesktop.loadComponentFromURL(sUrl, "_blank", 0, aProps())
	vT = ContexteTexte(oCalc, oPlage, oPlage.getDataArray(), aOpt, Nothing)
	ColorierDocument(oDoc, oPlage.getDataArray(), aPal, vT)
	oForme = InsererDansCalc(oDoc, oCalc, oPlage, NomFichier(sUrl), bImage)
	oDoc.close(True)
	InsererCarteT = CreerLien(oCalc, oSel, oForme, aPal(P_ID), sUrl, bImage, aOpt)
	If aOpt(O_LEGENDE) Or aOpt(O_VALEURS) Or aOpt(O_LIBELLES) Then RecalculerLien(oCalc, InsererCarteT)
End Function

' ---------------------------------------------------------------- textes
Function TextesT(oG As Object, sCode As String) As String
	Dim i As Long, nZ As Long, nT As Long, nL As Long, iPremierTexte As Long, iDerniereZone As Long
	Dim s As String, o As Object
	iPremierTexte = -1
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		If Left(o.Name, Len(PREFIXE_TEXTE)) = PREFIXE_TEXTE Then
			nT = nT + 1
			If iPremierTexte < 0 Then iPremierTexte = i
			If o.Name = PREFIXE_TEXTE & " " & sCode Then s = s & "texte '" & Replace(o.getString(), Chr(10), "/") _
				& "' couleur " & Hex(o.CharColor) & " centre " & (o.Position.X + o.Size.Width / 2) & "," _
				& (o.Position.Y + o.Size.Height / 2)
		ElseIf EstLegende(o.Name) Then
			nL = nL + 1
		Else
			nZ = nZ + 1
			iDerniereZone = i
			If o.Name = sCode Then s = s & " ; zone centre " & (o.Position.X + o.Size.Width / 2) & "," _
				& (o.Position.Y + o.Size.Height / 2) & " texte dans la zone='" & o.getString() & "'"
		End If
	Next i
	TextesT = nZ & " zones, " & nT & " textes, " & nL & " formes de légende ; textes après les zones : " _
		& (nT = 0 Or iPremierTexte > iDerniereZone) & " | " & sCode & " : " & s
End Function

Sub TestTextes()
	Dim oCalc As Object, oFeuille As Object, oG As Object, aOpt As Variant, sLien As String, i As Long, t As Long
	Dim aChamps As Variant, oDoc As Object, oPlage As Object, vT As Variant, aPals As Variant
	Dim aProps(0) As New com.sun.star.beans.PropertyValue
	OuvrirJournal()
	On Error GoTo ErrT
	ChargerTextes()
	oCalc = DocCalcT()
	oFeuille = oCalc.getSheets().getByIndex(0)
	aOpt = OptionsParDefaut()
	aOpt(O_VALEURS) = True : aOpt(O_LIBELLES) = True : aOpt(O_LEGENDE) = True
	aOpt(O_TAILLE_VALEURS) = 14 : aOpt(O_TAILLE_LIBELLES) = 7
	t = GetSystemTicks()
	sLien = InsererCarteT(aOpt, "jaune-rouge")
	oG = oFeuille.getDrawPage().getByIndex(0)
	Journal("TEST insertion (" & (GetSystemTicks() - t) & " ms) : " & TextesT(oG, "88"))
	' Modification d'une valeur : textes à jour, pas de doublons
	t = GetSystemTicks()
	For i = 1 To 200
		If oFeuille.getCellByPosition(0, i).getString() = "88" Then oFeuille.getCellByPosition(1, i).setValue(4321)
	Next i
	Journal("TEST modification (" & (GetSystemTicks() - t) & " ms) : " & TextesT(oG, "88"))
	' Carte d'une version précédente : texte dans la zone, retiré
	For i = 0 To oG.getCount() - 1
		If oG.getByIndex(i).Name = "88" Then oG.getByIndex(i).setString("ancien")
	Next i
	RecalculerLien(oCalc, sLien)
	Journal("TEST ancien texte dans la zone : " & TextesT(oG, "88"))
	' Valeurs et libellés désactivés : textes retirés
	aChamps = Split(oCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sLien), "|")
	aOpt(O_VALEURS) = False : aOpt(O_LIBELLES) = False
	oCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sLien, _
		ChampsLien(aChamps(1), aChamps(2), aChamps(3), False, aOpt))
	RecalculerLien(oCalc, sLien)
	Journal("TEST options désactivées : " & TextesT(oG, "88"))
	aOpt(O_VALEURS) = True : aOpt(O_LIBELLES) = True
	oCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sLien, _
		ChampsLien(aChamps(1), aChamps(2), aChamps(3), False, aOpt))
	RecalculerLien(oCalc, sLien)
	' Ouverture dans Draw
	aPals = ChargerPalettes()
	oPlage = PlageDeDonnees(PlageSelectionnee(oFeuille.getCellRangeByName("A1:C1048576")))
	aProps(0).Name = "Hidden" : aProps(0).Value = True
	oDoc = StarDesktop.loadComponentFromURL(ConvertToURL("/home/gabor/github/france-odg/output/departements-france.odg"), _
		"_blank", 0, aProps())
	vT = ContexteTexte(oCalc, oPlage, oPlage.getDataArray(), aOpt, oDoc)
	ColorierDocument(oDoc, oPlage.getDataArray(), aPals(IndicePalette(aPals, "jaune-rouge")), vT)
	Journal("TEST Draw : " & TextesT(oDoc.getDrawPages().getByIndex(0), "88"))
	oDoc.close(True)
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- éditeur de classes
Function EtatT() As String
	Dim k As Integer, s As String, aN As Variant
	aN = CompterClasses()
	For k = 0 To nClassesEd - 1
		s = s & "[" & NombreEd(aBornesEd(k)) & ".." & NombreEd(aBornesEd(k + 1)) & " " & Hex(aCouleursEd(k)) _
			& " n=" & aN(k) & "]"
	Next k
	EtatT = nClassesEd & " classes (champ " & oEditeur.getControl("numClasses").getText() & ") " & s _
		& " | borne choisie=" & iBorneEd & " champ='" & oEditeur.getControl("txtBorne").getText() & "'" _
		& " | erreur='" & oEditeur.getControl("lblErreur").getModel().Label & "'"
End Function

Sub SourisT(sQuoi As String, sNom As String, nX As Long)
	Dim oEv As Object
	oEv = CreateUnoStruct("com.sun.star.awt.MouseEvent")
	oEv.Source = oEditeur.getControl(sNom)
	oEv.X = nX
	oEv.Buttons = com.sun.star.awt.MouseButton.LEFT
	Select Case sQuoi
		Case "appui" : Barre_mousePressed(oEv)
		Case "glisse" : BarreMvt_mouseDragged(oEv)
		Case "relache" : Barre_mouseReleased(oEv)
	End Select
End Sub

Sub OuvrirEditeurT(sPalette As String)
	Dim aPals As Variant, oPlage As Object, s As String
	ChargerTextes()
	aPals = ChargerPalettes()
	oPlage = PlageDeDonnees(PlageSelectionnee(DocCalcT().getSheets().getByIndex(0).getCellRangeByName("A1:C1048576")))
	s = ConstruireEditeur("", aPals(IndicePalette(aPals, sPalette)), oPlage.getDataArray(), 0, 1, 2, Nothing)
End Sub

Sub TestEditeurComplet()
	Dim oFond As Object, oEv As Object, oNum As Object, k As Integer
	OuvrirJournal()
	On Error GoTo ErrT
	OuvrirEditeurT("bleus")
	oFond = oEditeur.getControl("fondBarre").getPosSize()
	oNum = oEditeur.getControl("numClasses")
	Journal("TEST depart : " & EtatT())
	oNum.up() : oNum.up()
	Journal("TEST 2 clics haut : " & EtatT())
	SourisT("appui", "pgn3", 1)
	SourisT("glisse", "fondBarre", oFond.Width / 2)
	SourisT("relache", "fondBarre", oFond.Width / 2)
	Journal("TEST borne 3 glissee au milieu : " & EtatT())
	SourisT("appui", "pgn3", 1)
	SourisT("glisse", "fondBarre", oFond.Width - 2)
	SourisT("relache", "fondBarre", oFond.Width - 2)
	Journal("TEST borne 3 glissee trop loin : " & EtatT())
	oEditeur.getControl("txtBorne").setText("450")
	Journal("TEST borne 3 saisie 450 : " & EtatT())
	oEditeur.getControl("txtBorne").setText("5000")
	Journal("TEST borne 3 saisie 5000 : " & EtatT())
	aCouleursEd(2) = RGB(255, 0, 0)
	oEv = CreateUnoStruct("com.sun.star.awt.ActionEvent")
	oEv.Source = oEditeur.getControl("btnLineaire")
	EditeurBouton_actionPerformed(oEv)
	Journal("TEST lineaire (7 classes et rouge gardes) : " & EtatT())
	oNum.setText("1")
	Journal("TEST saisie '1' (ignoree) : " & EtatT())
	oNum.setText("12")
	Journal("TEST saisie '12' : " & EtatT())
	For k = 1 To 4 : oNum.down() : Next k
	Journal("TEST 4 clics bas : " & EtatT())
	For k = 1 To 30 : oNum.up() : Next k
	Journal("TEST 30 clics haut (max 20) : " & EtatT())
	SourisT("appui", "seg1", 3)
	SourisT("relache", "seg1", 3)
	Journal("TEST clic sur une couleur : " & EtatT())
	oEditeur.dispose()
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- lisibilité
Sub EssaiLisibiliteT(oDoc As Object, oPage As Object, nX As Long, nY As Long, sMethode As String, nTexte As Long, nFond As Long)
	Dim oFond As Object, oT As Object, k As Integer, aDx As Variant, aDy As Variant
	Dim oPos As New com.sun.star.awt.Point, oTaille As New com.sun.star.awt.Size
	oFond = oDoc.createInstance("com.sun.star.drawing.RectangleShape")
	oPage.add(oFond)
	oPos.X = nX : oPos.Y = nY : oTaille.Width = 2600 : oTaille.Height = 1200
	oFond.setPosition(oPos) : oFond.setSize(oTaille)
	oFond.FillColor = nFond
	If sMethode = "copies" Then
		aDx = Array(-25, 25, -25, 25, 0, 0, -35, 35)
		aDy = Array(-25, -25, 25, 25, -35, 35, 0, 0)
		For k = 0 To 7
			oT = TexteEssaiT(oDoc, oPage, nX + 300 + aDx(k), nY + 300 + aDy(k), IIf(nTexte = 0, RGB(255, 255, 255), 0))
		Next k
	End If
	oT = TexteEssaiT(oDoc, oPage, nX + 300, nY + 300, nTexte)
	If sMethode = "lueur" Then
		oT.GlowEffectRadius = 50
		oT.GlowEffectColor = IIf(nTexte = 0, RGB(255, 255, 255), 0)
	ElseIf sMethode = "etiquette" Then
		oT.FillStyle = com.sun.star.drawing.FillStyle.SOLID
		oT.FillColor = RGB(255, 255, 255)
		oT.FillTransparence = 25
		oT.CharColor = 0
		oT.TextLeftDistance = 60 : oT.TextRightDistance = 60
	End If
End Sub

Function TexteEssaiT(oDoc As Object, oPage As Object, nX As Long, nY As Long, nCouleur As Long) As Object
	Dim oT As Object
	Dim oPos As New com.sun.star.awt.Point
	oT = oDoc.createInstance("com.sun.star.drawing.TextShape")
	oPage.add(oT)
	oT.TextAutoGrowWidth = True : oT.TextAutoGrowHeight = True
	oT.setString("Vosges" & Chr(10) & "946")
	oT.CharHeight = 12 : oT.CharFontName = "Liberation Sans" : oT.CharColor = nCouleur
	oPos.X = nX : oPos.Y = nY
	oT.setPosition(oPos)
	TexteEssaiT = oT
End Function

' Compare les méthodes de lisibilité sur un fond moyen (orange) et un fond
' de même teinte que le texte : image dans essai.png
Sub TestLisibilite()
	Dim oDoc As Object, oPage As Object, i As Integer, sMethode As Variant, aProps(0) As New com.sun.star.beans.PropertyValue
	OuvrirJournal()
	On Error GoTo ErrT
	aProps(0).Name = "Hidden" : aProps(0).Value = True
	oDoc = StarDesktop.loadComponentFromURL("private:factory/sdraw", "_blank", 0, aProps())
	oPage = oDoc.getDrawPages().getByIndex(0)
	For Each sMethode In Array("aucune", "lueur", "copies", "etiquette")
		EssaiLisibiliteT(oDoc, oPage, 1000 + i * 3000, 1000, sMethode, 0, RGB(253, 141, 60))
		EssaiLisibiliteT(oDoc, oPage, 1000 + i * 3000, 2500, sMethode, 0, RGB(90, 90, 90))
		EssaiLisibiliteT(oDoc, oPage, 1000 + i * 3000, 4000, sMethode, RGB(255, 255, 255), RGB(253, 141, 60))
		EssaiLisibiliteT(oDoc, oPage, 1000 + i * 3000, 5500, sMethode, RGB(255, 255, 255), RGB(230, 230, 230))
		i = i + 1
	Next sMethode
	ExporterPageT(oPage, "essai.png")
	Journal("TEST image essai.png creee")
	oDoc.close(True)
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

Sub ExporterPageT(oPage As Object, sNom As String)
	Dim oFiltre As Object, sDossier As String
	Dim aArgs(2) As New com.sun.star.beans.PropertyValue, aDonnees(0) As New com.sun.star.beans.PropertyValue
	sDossier = Environ("CALCMAP_TESTS")
	aDonnees(0).Name = "PixelWidth" : aDonnees(0).Value = 1400
	aArgs(0).Name = "URL" : aArgs(0).Value = ConvertToURL(sDossier & "/" & sNom)
	aArgs(1).Name = "MediaType" : aArgs(1).Value = "image/png"
	aArgs(2).Name = "FilterData" : aArgs(2).Value = aDonnees()
	oFiltre = CreateUnoService("com.sun.star.drawing.GraphicExportFilter")
	oFiltre.setSourceDocument(oPage)
	oFiltre.filter(aArgs())
End Sub

Sub TestProprietesLueur()
	Dim oCalc As Object, oT As Object, oInfo As Object, sNom As Variant, s As String
	OuvrirJournal()
	On Error GoTo ErrT
	oCalc = DocCalcT()
	oT = oCalc.createInstance("com.sun.star.drawing.TextShape")
	oCalc.getSheets().getByIndex(0).getDrawPage().add(oT)
	oInfo = oT.getPropertySetInfo()
	For Each sNom In Array("GlowTextRadius", "GlowTextColor", "GlowTextTransparency", "GlowEffectRadius", "CharContoured", "CharShadowed")
		s = s & sNom & "=" & oInfo.hasPropertyByName(sNom) & " "
	Next sNom
	Journal("TEST proprietes : " & s)
	oT.GlowTextRadius = 60
	oT.GlowTextColor = RGB(255, 255, 255)
	Journal("TEST relues : rayon=" & oT.GlowTextRadius & " couleur=" & Hex(oT.GlowTextColor))
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

' Carte des départements avec chaque réglage de lisibilité : nombre de formes,
' temps de recoloriage, image lisibilite-<réglage>.png
Sub TestLisibiliteCarte()
	Dim oCalc As Object, oFeuille As Object, aOpt As Variant, sLien As String, sMode As Variant, t As Long
	Dim oG As Object, i As Long, aChamps As Variant
	OuvrirJournal()
	On Error GoTo ErrT
	ChargerTextes()
	oCalc = DocCalcT()
	oFeuille = oCalc.getSheets().getByIndex(0)
	aOpt = OptionsParDefaut()
	Journal("TEST lisibilite par defaut : " & aOpt(O_LISIBILITE))
	aOpt(O_VALEURS) = True : aOpt(O_LIBELLES) = True : aOpt(O_TAILLE_LIBELLES) = 7
	sLien = InsererCarteT(aOpt, "jaune-rouge-10")
	oG = oFeuille.getDrawPage().getByIndex(0)
	aChamps = Split(oCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sLien), "|")
	For Each sMode In Array("aucune", "contour", "etiquette")
		aOpt(O_LISIBILITE) = sMode
		oCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sLien, _
			ChampsLien(aChamps(1), aChamps(2), aChamps(3), False, aOpt))
		t = GetSystemTicks()
		RecalculerLien(oCalc, sLien)
		Journal("TEST " & sMode & " : " & oG.getCount() & " formes dans la carte, recoloriage " & (GetSystemTicks() - t) & " ms")
		ExporterFormeT(oG, "lisibilite-" & sMode & ".png")
	Next sMode
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

Sub ExporterFormeT(oForme As Object, sNom As String)
	Dim oFiltre As Object
	Dim aArgs(2) As New com.sun.star.beans.PropertyValue, aDonnees(0) As New com.sun.star.beans.PropertyValue
	aDonnees(0).Name = "PixelWidth" : aDonnees(0).Value = 2000
	aArgs(0).Name = "URL" : aArgs(0).Value = ConvertToURL(Environ("CALCMAP_TESTS") & "/" & sNom)
	aArgs(1).Name = "MediaType" : aArgs(1).Value = "image/png"
	aArgs(2).Name = "FilterData" : aArgs(2).Value = aDonnees()
	oFiltre = CreateUnoService("com.sun.star.drawing.GraphicExportFilter")
	oFiltre.setSourceDocument(oForme)
	oFiltre.filter(aArgs())
End Sub

' ---------------------------------------------------------------- dialogue principal
' Le dialogue ne s'affiche pas sans écran : il est construit, puis ses
' réglages relus ; ils doivent revenir à l'identique
Sub TestDialogue()
	Dim aOpt As Variant, aPals As Variant, iC As Integer, iP As Integer, bI As Boolean, n As Integer, sAvant As String
	Dim oPlage As Object, sMode As Variant
	OuvrirJournal()
	On Error GoTo ErrT
	ChargerTextes()
	aPals = ChargerPalettes()
	oPlage = PlageDeDonnees(PlageSelectionnee(DocCalcT().getSheets().getByIndex(0).getCellRangeByName("A1:C1048576")))
	aDonneesDialogue = oPlage.getDataArray()
	For Each sMode In Array("aucune", "contour", "etiquette")
		aOpt = OptionsParDefaut()
		aOpt(O_VALEURS) = True : aOpt(O_STYLE_LIBELLES) = "BI" : aOpt(O_TAILLE_LIBELLES) = 13
		aOpt(O_LISIBILITE) = sMode
		sAvant = TexteOptions(aOpt)
		bI = True
		n = DialogueCartes(Array(NouvelleCarte("departements", "Dep", "d", "", "d.odg", 1, "", True)), aPals, "etat", _
			"", "bleus", iC, iP, bI, aOpt, NomsColonnes(oPlage), Nothing, False)
		Journal("TEST dialogue " & sMode & " : relu a l'identique=" & (TexteOptions(aOpt) = sAvant) & " ; " & TexteOptions(aOpt))
	Next sMode
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub
