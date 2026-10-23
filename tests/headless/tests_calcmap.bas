
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
	RecalculerLien(oCalc, sLien, True)
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
	Journal("TEST contour par defaut : valeurs " & aOpt(O_CONTOUR_VALEURS) & ", libelles " & aOpt(O_CONTOUR_LIBELLES))
	aOpt(O_VALEURS) = True : aOpt(O_LIBELLES) = True : aOpt(O_TAILLE_LIBELLES) = 7
	sLien = InsererCarteT(aOpt, "jaune-rouge-10")
	oG = oFeuille.getDrawPage().getByIndex(0)
	aChamps = Split(oCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sLien), "|")
	For Each sMode In Array("aucune", "contour", "valeurs seules")
		aOpt(O_CONTOUR_VALEURS) = IIf(sMode = "aucune", "aucune", "contour")
		aOpt(O_CONTOUR_LIBELLES) = IIf(sMode = "contour", "contour", "aucune")
		oCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sLien, _
			ChampsLien(aChamps(1), aChamps(2), aChamps(3), False, aOpt))
		t = GetSystemTicks()
		RecalculerLien(oCalc, sLien)
		Journal("TEST " & sMode & " : " & oG.getCount() & " formes dans la carte, recoloriage " & (GetSystemTicks() - t) & " ms ; " & CopieT(oG))
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
	For Each sMode In Array("aucune", "contour")
		aOpt = OptionsParDefaut()
		aOpt(O_VALEURS) = True : aOpt(O_STYLE_LIBELLES) = "BI" : aOpt(O_TAILLE_LIBELLES) = 13
		aOpt(O_CONTOUR_VALEURS) = sMode : aOpt(O_CONTOUR_LIBELLES) = IIf(sMode = "aucune", "contour", "aucune")
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

' ---------------------------------------------------------------- mise à jour partielle
' État de la carte : couleur, texte alternatif et textes (triés) de chaque zone
Function EtatCarteT(oG As Object) As String
	Dim i As Long, o As Object, aT() As String, n As Long
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		ReDim Preserve aT(n) As String
		If EstLegende(o.Name) Then
			aT(n) = "L " & o.getShapeType() & " " & o.Position.X & "," & o.Position.Y & IIf(o.getShapeType() = "com.sun.star.drawing.TextShape", " '" & o.getString() & "'", " " & Hex(o.FillColor))
		ElseIf Left(o.Name, Len(PREFIXE_TEXTE)) = PREFIXE_TEXTE Then
			aT(n) = "T " & o.Name & " '" & Replace(o.getString(), Chr(10), "/") & "' " & Hex(o.CharColor) & " " & o.Position.X & "," & o.Position.Y
		Else
			aT(n) = "Z " & o.Name & " " & Hex(o.FillColor) & " [" & o.Description & "]"
		End If
		n = n + 1
	Next i
	TrierTextes(aT)
	EtatCarteT = Join(aT, Chr(10))
End Function

' Modifie des cellules (adresse=valeur ; vide : cellule effacée) puis compare
' la mise à jour (partielle ou non) à un recalcul complet
Sub CasPartielT(sCas As String, aModifs As Variant, sLien As String, oG As Object)
	Dim oCalc As Object, oFeuille As Object, k As Integer, a As Variant, t As Long, sApres As String, sComplet As String
	Dim nZones As Long
	oCalc = DocCalcT()
	oFeuille = oCalc.getSheets().getByIndex(0)
	t = GetSystemTicks()
	For k = 0 To UBound(aModifs)
		a = Split(aModifs(k), "=")
		If a(1) = "" Then
			oFeuille.getCellRangeByName(a(0)).setString("")
		ElseIf IsNumeric(a(1)) Then
			oFeuille.getCellRangeByName(a(0)).setValue(CDbl(a(1)))
		Else
			oFeuille.getCellRangeByName(a(0)).setString(a(1))
		End If
	Next k
	t = GetSystemTicks() - t
	nZones = nDerniereMaj
	sApres = EtatCarteT(oG)
	RecalculerLien(oCalc, sLien, True)
	sComplet = EtatCarteT(oG)
	Journal("TEST " & sCas & " : " & IIf(nZones < 0, "carte entiere", nZones & " zone(s)") & " en " & t & " ms ; identique au recalcul complet : " & (sApres = sComplet))
	If sApres <> sComplet Then Journal("TEST   difference : " & DifferenceT(sApres, sComplet))
End Sub

Function DifferenceT(a As String, b As String) As String
	Dim aa As Variant, bb As Variant, i As Long
	aa = Split(a, Chr(10)) : bb = Split(b, Chr(10))
	For i = 0 To UBound(aa)
		If i > UBound(bb) Then DifferenceT = "en trop : " & aa(i) : Exit Function
		If aa(i) <> bb(i) Then DifferenceT = aa(i) & "  <>  " & bb(i) : Exit Function
	Next i
	DifferenceT = "nombre de formes " & (UBound(aa) + 1) & " / " & (UBound(bb) + 1)
End Function

Sub TestPartiel()
	Dim oCalc As Object, oFeuille As Object, aOpt As Variant, sLien As String, oG As Object, i As Long, aModifs() As String
	OuvrirJournal()
	On Error GoTo ErrT
	ChargerTextes()
	oCalc = DocCalcT()
	oFeuille = oCalc.getSheets().getByIndex(0)
	aOpt = OptionsParDefaut()
	aOpt(O_VALEURS) = True : aOpt(O_LIBELLES) = True : aOpt(O_LEGENDE) = True
	sLien = InsererCarteT(aOpt, "jaune-rouge")
	oG = oFeuille.getDrawPage().getByIndex(0)
	' Lignes : A2 = 01 (Ain, 389) ... ; 88 en ligne 90 ; min 10 (Lozère, 48), max 988 (Gard, 30)
	CasPartielT("valeur changee dans les bornes (Ain 389 -> 700)", Array("B2=700"), sLien, oG)
	CasPartielT("meme valeur ressaisie", Array("B2=700"), sLien, oG)
	CasPartielT("libelle change (Ain -> AIN)", Array("C2=AIN"), sLien, oG)
	CasPartielT("code change (01 -> 02, doublon)", Array("A2=02"), sLien, oG)
	CasPartielT("code retabli", Array("A2=01"), sLien, oG)
	CasPartielT("nouveau maximum (Ain -> 2000)", Array("B2=2000"), sLien, oG)
	CasPartielT("valeur effacee (Ain)", Array("B2="), sLien, oG)
	CasPartielT("valeur remise (Ain 389)", Array("B2=389"), sLien, oG)
	ReDim aModifs(24) As String
	For i = 0 To 24
		aModifs(i) = "B" & (i + 3) & "=" & (500 + i)
	Next i
	CasPartielT("25 valeurs modifiees une a une", aModifs, sLien, oG)
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

' Paragraphes d'une copie de contour du texte des Vosges : texte et transparence
Function CopieT(oG As Object) As String
	Dim i As Long, o As Object, oEnum As Object, p As Object, s As String
	For i = 0 To oG.getCount() - 1
		o = oG.getByIndex(i)
		If o.Name = PREFIXE_TEXTE & " 88" Then
			oEnum = o.getText().createEnumeration()
			s = ""
			Do While oEnum.hasMoreElements()
				p = oEnum.nextElement()
				s = s & "[" & p.getString() & "]"
			Loop
			CopieT = "1re forme texte des Vosges : " & s & " couleur " & Hex(o.CharColor)
			Exit Function
		End If
	Next i
End Function

' ---------------------------------------------------------------- aperçus
' Une image SVG s'affiche-t-elle dans un contrôle image de dialogue ?
Sub TestApercuSvg()
	Dim oModele As Object, oDlg As Object, oImg As Object, oGraphique As Object
	Dim aSource(0) As New com.sun.star.beans.PropertyValue
	OuvrirJournal()
	On Error GoTo ErrT
	aSource(0).Name = "URL" : aSource(0).Value = ConvertToURL(Environ("CALCMAP_TESTS") & "/../essai.svg")
	oGraphique = CreateUnoService("com.sun.star.graphic.GraphicProvider").queryGraphic(aSource())
	Journal("TEST svg charge : " & Not IsNull(oGraphique) & " ; type " & oGraphique.GraphicType & " ; taille " & oGraphique.Size100thMM.Width & "x" & oGraphique.Size100thMM.Height)
	oModele = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
	oModele.Width = 200 : oModele.Height = 150
	AjouterControle(oModele, "ImageControl", "img", 10, 10, 100, 80, Array("ScaleImage", "ScaleMode", "ImageURL"), _
		Array(True, com.sun.star.awt.ImageScaleMode.ISOTROPIC, aSource(0).Value))
	oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
	oDlg.setModel(oModele)
	oDlg.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Nothing)
	Journal("TEST controle image : graphique present=" & Not IsNull(oDlg.getControl("img").getModel().Graphic))
	oDlg.dispose()
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- catalogue
Sub ParametreT(sNom As String, v As Variant)
	Dim oAcces As Object, aArgs(0) As New com.sun.star.beans.PropertyValue
	aArgs(0).Name = "nodepath" : aArgs(0).Value = NOEUD_PARAMETRES
	oAcces = CreateUnoService("com.sun.star.configuration.ConfigurationProvider") _
		.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", aArgs())
	oAcces.setPropertyValue(sNom, v)
	oAcces.commitChanges()
End Sub

Function ListeT(aCartes As Variant) As String
	Dim i As Long, s As String
	For i = 0 To UBound(aCartes)
		s = s & aCartes(i)(C_ID) & "[" & aCartes(i)(C_SOURCE) & ", " & aCartes(i)(C_OBJETS) & " zones" _
			& IIf(aCartes(i)(C_APERCU) <> "", ", apercu", "") & IIf(aCartes(i)(C_CODES) <> "", ", codes " & aCartes(i)(C_CODES), "") & "] "
	Next i
	ListeT = (UBound(aCartes) + 1) & " carte(s) : " & s
End Function

Sub TestCatalogue()
	Dim sTests As String, sFixtures As String, sFrance As String, s As String, aCartes As Variant, b As Boolean, i As Long
	OuvrirJournal()
	On Error GoTo ErrT
	ChargerTextes()
	Journal("TEST parametres par defaut : catalogue=" & UrlCatalogue() & " ; zones max=" & ZonesMax())
	sTests = ConvertToURL(Environ("CALCMAP_TESTS"))
	sFixtures = ConvertToURL(Environ("CALCMAP_FIXTURES"))
	sFrance = ConvertToURL("/home/gabor/github/france-odg/output")
	' Source france-odg : le catalogue du projet, aperçus en fichiers locaux
	s = Replace(LireTexte(sFrance & "/maps.json"), "https://raw.githubusercontent.com/gaborjager/france-odg/main/output", sFrance)
	EcrireTexte(sTests & "/source-france.json", s)
	EcrireTexte(sTests & "/source-entreprise.json", Replace(Replace(LireTexte(sFixtures & "/source-entreprise.json"), _
		"@TESTS@", sTests), "@FRANCEODG@", sFrance))
	EcrireTexte(sTests & "/agregat.json", Replace(LireTexte(sFixtures & "/agregat.json"), "@TESTS@", sTests))
	ParametreT("Catalogue", sTests & "/agregat.json")
	ParametreT("ZonesMax", 200)
	aCartes = ChargerCatalogue(DossierCartes(), b)
	Journal("TEST max 200 : en ligne=" & b & " ; " & ListeT(aCartes))
	ParametreT("ZonesMax", 3000)
	aCartes = ChargerCatalogue(DossierCartes(), b)
	Journal("TEST max 3000 : " & ListeT(aCartes))
	For i = 0 To UBound(aCartes)
		If aCartes(i)(C_ID) = "departements" Or aCartes(i)(C_ID) = "secteurs" Then
			s = ApercuLocal(aCartes(i))
			Journal("TEST apercu " & aCartes(i)(C_ID) & " : '" & s & "' lisible=" & Not IsNull(GraphiqueDe(s)))
		End If
	Next i
	Journal("TEST image par defaut lisible : " & Not IsNull(GraphiqueDe(DossierExtension() & "icons/nopreview.svg")))
	' Catalogue injoignable : copie en cache
	ParametreT("Catalogue", sTests & "/absent.json")
	aCartes = ChargerCatalogue(DossierCartes(), b)
	Journal("TEST catalogue injoignable : en ligne=" & b & " ; " & ListeT(aCartes))
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- paramètres
' Paramètres lus, puis modifiés comme dans la configuration avancée
Sub TestParametres()
	OuvrirJournal()
	Journal("TEST parametres lus : catalogue=" & UrlCatalogue() & " ; zones max=" & ZonesMax())
	On Error GoTo Refus
	ParametreT("ZonesMax", 7)
	Journal("TEST modification par l'utilisateur acceptee : zones max=" & ZonesMax())
	FermerJournal()
	Exit Sub
Refus:
	Journal("TEST modification par l'utilisateur refusee (" & Error$ & ") : zones max=" & ZonesMax())
	FermerJournal()
End Sub

' ---------------------------------------------------------------- icône sur une carte
' Carte sélectionnée (ou une de ses zones) : l'icône ouvre sa configuration
' (le dialogue ne s'affiche pas sans écran) et n'insère pas de nouvelle carte
Sub TestIconeSurCarte()
	Dim oCalc As Object, oFeuille As Object, aOpt As Variant, sLien As String, oG As Object, n As Long
	OuvrirJournal()
	On Error GoTo ErrT
	ChargerTextes()
	oCalc = DocCalcT()
	oFeuille = oCalc.getSheets().getByIndex(0)
	aOpt = OptionsParDefaut()
	sLien = InsererCarteT(aOpt, "bleus")
	oG = oFeuille.getDrawPage().getByIndex(0)
	FermerJournal()
	oCalc.getCurrentController().select(oG)
	CalcMap()
	n = oFeuille.getDrawPage().getCount()
	oCalc.getCurrentController().select(oG.getByIndex(10))
	CalcMap()
	JournalAjout("TEST apres l'icone sur la carte puis sur une zone : " & oFeuille.getDrawPage().getCount() & " forme(s) dans la feuille (avant : " & n & ")")
	Exit Sub
ErrT:
	JournalAjout("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub


' ---------------------------------------------------------------- traductions
' Chaque langue : fichier chargé, toutes les clés traduites, exemple lu en
' UTF-8 ; dialogue construit dans quelques langues
Sub TestTraductions()
	Dim aCodes As Variant, sCode As Variant, aCles As Variant, sCle As Variant, nManque As Integer, s As String
	Dim oPlage As Object, aOpt As Variant, aPals As Variant, iC As Integer, iP As Integer, bI As Boolean, n As Integer
	OuvrirJournal()
	On Error GoTo ErrT
	aCodes = Split("en fr ar bg cs da de el es et fi ga hi hr hu id it ja ko lt lv mt nb nl pl pt ro ru sk sl sv tr uk zh", " ")
	aCles = Array("dlg.grp_map", "dlg.show_values", "ed.help", "msg.inserted_as_image", "dlg.zones", "menu.configure")
	For Each sCode In aCodes
		sLangue = sCode
		ChargerTextes()
		nManque = 0
		For Each sCle In aCles
			If Tr(sCle) = sCle Then nManque = nManque + 1
		Next sCle
		s = s & sCode & "=" & Tr("dlg.grp_map") & IIf(nManque > 0, "(!" & nManque & ")", "") & " "
	Next sCode
	Journal("TEST langues : " & s)
	sLangue = "fr" : ChargerTextes()
	Journal("TEST francais different de l'anglais : " & Tr("dlg.show_values") & " / %3 : " & Tr("ed.info", 1, 2, 96))
	aPals = ChargerPalettes()
	oPlage = PlageDeDonnees(PlageSelectionnee(DocCalcT().getSheets().getByIndex(0).getCellRangeByName("A1:C1048576")))
	aDonneesDialogue = oPlage.getDataArray()
	For Each sCode In Array("de", "ar", "ja")
		sLangue = sCode : ChargerTextes()
		aOpt = OptionsParDefaut()
		bI = True
		n = DialogueCartes(Array(NouvelleCarte("departements", "Dep", "d", "", "d.odg", 1, "", True)), aPals, "", _
			"", "bleus", iC, iP, bI, aOpt, NomsColonnes(oPlage), Nothing, False)
		Journal("TEST dialogue en " & sCode & " : construit (resultat " & n & ")")
	Next sCode
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub

' ---------------------------------------------------------------- colonnes
' Noms des colonnes : 3 premières valeurs, « , ... » seulement si d'autres suivent
Sub TestNomsColonnes()
	Dim oFeuille As Object
	OuvrirJournal()
	On Error GoTo ErrT
	oFeuille = DocCalcT().getSheets().getByIndex(0)
	Journal("TEST 96 lignes : " & Join(NomsColonnes(PlageDeDonnees(PlageSelectionnee(oFeuille.getCellRangeByName("A1:C1048576")))), " | "))
	Journal("TEST 3 lignes : " & Join(NomsColonnes(oFeuille.getCellRangeByName("A1:C4")), " | "))
	Journal("TEST 2 lignes : " & Join(NomsColonnes(oFeuille.getCellRangeByName("A1:C3")), " | "))
	FermerJournal()
	Exit Sub
ErrT:
	Journal("TEST ERREUR ligne " & Erl & " : " & Error$)
	Resume Next
End Sub
