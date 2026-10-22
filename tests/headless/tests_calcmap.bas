
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
