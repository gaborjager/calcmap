REM  CalcMap — cartes thématiques pour LibreOffice Calc
REM  ===================================================
REM  Colorie les zones d'un fond de carte .odg (une forme par zone, nommée
REM  avec son code, ex. code INSEE) à partir de 2 colonnes Calc : code | valeur.
REM
REM  Utilisation :
REM   1. Dans Calc, sélectionner les données : par défaut code en 1re colonne,
REM      valeur en 2e, libellé facultatif en 3e ; le dialogue permet de choisir
REM      d'autres colonnes de la sélection.
REM   2. Cliquer sur l'icône globe (ou Outils > CalcMap : colorier une carte…).
REM   3. Choisir le fond de carte dans la liste (catalogue france-odg), ou un
REM      autre fichier .odg avec « Autre fichier… », et la palette de couleurs.
REM
REM  Fonds de carte : le catalogue maps.json de france-odg est lu à chaque
REM  exécution. Les cartes sont téléchargées dans le profil LibreOffice
REM  ($(user)/calcmap/) et retéléchargées seulement si leur empreinte
REM  SHA-256 diffère de celle du catalogue. Sans réseau, seules les cartes
REM  déjà téléchargées sont proposées.
REM
REM  Insertion dans Calc (option du dialogue) : la carte coloriée est placée
REM  dans la feuille, à droite des données. Jusqu'à SEUIL_FORMES_CALC zones,
REM  les formes sont recréées dans la feuille (chaque zone garde son nom et
REM  son texte alternatif) ; au-delà (communes), la carte est insérée en image.
REM
REM  Mise à jour automatique : une carte insérée dans Calc reste liée à ses
REM  données (plage nommée CalcMap_N + propriété du document du même
REM  nom). Toute modification des valeurs recolore la carte. Le lien est
REM  enregistré avec le classeur et reconnecté à son ouverture (événement
REM  « Ouvrir le document » déclaré par l'extension dans Events.xcu).
REM
REM  Valeurs et libellés sur la carte (options du dialogue) : le libellé
REM  (3e colonne) et la valeur de chaque zone sont écrits comme texte de la
REM  forme, sur deux paragraphes, la valeur au format de nombre de la colonne
REM  Calc, en blanc sur les couleurs foncées et en noir sur les claires, aux
REM  tailles choisies dans le dialogue (points, sur la carte insérée dans Calc ;
REM  même aspect dans Draw). Au-delà de SEUIL_FORMES_CALC
REM  zones, ni valeurs ni libellés ne sont affichés.
REM
REM  Légende (option) : une case de couleur par classe, avec ou sans ses bornes,
REM  à droite de la carte et dans son groupe (formes nommées PREFIXE_LEGENDE),
REM  recréée à chaque recoloriage.
REM
REM  Configuration : clic droit sur une carte insérée -> « Configuration de la
REM  carte… » (menu contextuel intercepté, IntercepterMenu) rouvre le dialogue
REM  pour changer palette, valeurs et libellés de cette carte.
REM
REM  Palettes : palettes.json de l'extension, complété (ou remplacé, à id
REM  égal) par $(user)/calcmap/palettes.json. Le nombre de couleurs
REM  d'une palette fixe le nombre de classes ; une palette "continuous"
REM  donne un dégradé continu (couleur interpolée selon la valeur exacte).
REM
REM  Langues : i18n/messages_<langue>.properties de l'extension, selon la
REM  langue de l'interface ; l'anglais sert de repli clé par clé. Le journal
REM  reste en français.
REM
REM  Robustesse : la macro ne doit jamais afficher d'erreur Basic.
REM   - lignes vides, en-têtes, codes vides : ignorés
REM   - valeur absente, vide, texte non numérique, cellule en erreur :
REM     la zone garde la couleur par défaut (gris)
REM   - nombres saisis comme du texte ("1 234,5") : acceptés
REM   - zéros initiaux perdus par Calc (01 -> 1) : retrouvés
REM
REM  Journal : chaque exécution écrase le fichier calcmap.log dans le
REM  profil LibreOffice (Windows : %APPDATA%\LibreOffice\4\user\,
REM  Linux : ~/.config/libreoffice/4/user/).
REM  Mettre JOURNAL_DETAILLE à True pour tracer aussi chaque zone.

Option Explicit

Const VERSION = "0.5.0"
Const ID_EXTENSION = "fr.github.gaborjager.calcmap"
Const JOURNAL_DETAILLE = False      ' True : une ligne par zone dans le journal
Const ECRIRE_TEXTE_ALT = True       ' True : valeur ajoutée au texte alternatif (Description)
Const NOM_JOURNAL = "calcmap.log"
Const URL_CATALOGUE = "https://raw.githubusercontent.com/gaborjager/france-odg/refs/heads/main/output/maps.json"
Const DOSSIER_CARTES = "calcmap"         ' cache des fonds de carte, dans le profil
Const ANCIEN_DOSSIER = "coloriercarte"   ' cache de l'ancienne extension ColorierCarte (repris)
Const NOM_CATALOGUE = "maps.json"        ' copie locale du catalogue (mode hors ligne)
Const CARTE_PAR_DEFAUT = "departements"
Const CARTES_IGNOREES = "|cantons|communes|"   ' cartes du catalogue non proposées (id entre |) : trop lourdes
Const NOM_PALETTES = "palettes.json"     ' palettes de l'extension et de l'utilisateur
Const NOM_PREFERENCES = "preferences.properties"   ' dernière carte et palette choisies
Const PALETTE_PAR_DEFAUT = "bleus"
Const NB_NUANCES_APERCU = 12             ' cases d'aperçu de palette dans le dialogue
Const SEUIL_FORMES_CALC = 5000           ' au-delà, carte insérée dans Calc en image
Const LARGEUR_CARTE_CALC = 16000         ' largeur de la carte insérée (1/100 mm)
Const POLICE_TEXTE = "Liberation Sans"   ' valeurs et libellés sur la carte (fournie avec LibreOffice)
Const TAILLE_TEXTE = 10                  ' taille par défaut des valeurs et libellés (points)
Const TAILLE_TEXTE_MIN = 2
Const TAILLE_TEXTE_MAX = 72
Const LARGEUR_IMAGE_PX = 2400            ' largeur de l'image PNG (cartes détaillées)
Const NOM_IMAGE = "carte-inseree.png"    ' image temporaire, dans le cache
Const URL_CONFIGURER = "vnd.sun.star.script:CalcMap.CalcMap.ConfigurerCarte?language=Basic&location=application"
Const PREFIXE_LEGENDE = "CalcMap légende"   ' nom des formes de la légende
Const NB_COLONNES_MAX = 26               ' colonnes proposées au plus dans le dialogue
Const PREFIXE_LIEN = "CalcMap_"          ' plages nommées et propriétés des cartes liées
Const ANCIEN_PREFIXE_LIEN = "ColorierCarte_"   ' liens créés par l'ancienne extension (reconnus)

' Champs d'une carte du catalogue (tableau renvoyé par NouvelleCarte)
Const C_ID = 0
Const C_TITRE = 1
Const C_DESCRIPTION = 2
Const C_URL = 3
Const C_FICHIER = 4
Const C_TAILLE = 5
Const C_SHA256 = 6
Const C_LOCALE = 7

' Options d'une carte (tableau renvoyé par OptionsParDefaut). Même ordre dans
' le lien (à partir du 6e champ) ; noms dans les préférences : NOMS_OPTIONS.
Const O_VALEURS = 0
Const O_TAILLE_VALEURS = 1
Const O_LIBELLES = 2
Const O_TAILLE_LIBELLES = 3
Const O_POLICE_VALEURS = 4
Const O_STYLE_VALEURS = 5           ' lettres B (gras), I (italique), U (souligné)
Const O_POLICE_LIBELLES = 6
Const O_STYLE_LIBELLES = 7
Const O_COL_CODE = 8                ' colonnes de la plage (0 : 1re) ; propres à chaque carte
Const O_COL_VALEUR = 9
Const O_COL_LIBELLE = 10
Const O_LEGENDE = 11
Const O_LEGENDE_VALEURS = 12
Const NB_OPTIONS = 13
Const NOMS_OPTIONS = "valeurs|taille_valeurs|libelles|taille_libelles|police_valeurs|style_valeurs|police_libelles|style_libelles|col_code|col_valeur|col_libelle|legende|legende_valeurs"
Const TYPES_OPTIONS = "bnbnsssscccbb"    ' b : oui/non, n : taille, s : texte, c : colonne

' Contexte d'un coloriage (tableau renvoyé par ContexteTexte)
Const T_FORMATEUR = 0
Const T_FORMAT = 1
Const T_OPTIONS = 2
Const T_DOCUMENT = 3                ' document où créer la légende (Nothing : pas de légende)
Const T_TITRE = 4                   ' titre de la légende : en-tête de la colonne des valeurs

' Champs d'une palette (tableau renvoyé par NouvellePalette)
Const P_ID = 0
Const P_NOM = 1
Const P_COULEURS = 2
Const P_VIDE = 3
Const P_CONTINU = 4                 ' True : dégradé continu, sinon classes

Private oJournal As Object          ' flux texte du journal (Null si indisponible)
Private sCheminJournal As String
Private oDialogue As Object         ' dialogue de choix de carte (pour ses écouteurs)
Private aCartesDialogue As Variant
Private aPalettesDialogue As Variant
Private oTextes As Object           ' textes traduits (clé -> texte), chargés au premier usage
Private sLangue As String
' Cartes liées : Global et non Private, car ces variables doivent survivre à la
' fin de la macro (l'écouteur de modification est appelé plus tard)
Global aLiens() As Variant          ' cartes liées suivies : Array(document, plage, écouteur, nom)
Global nLiens As Long
Global bRecalculEnCours As Boolean
Global aMenus() As Variant          ' menus contextuels interceptés : Array(contrôleur, intercepteur)
Global nMenus As Long
Private sJson As String             ' état du lecteur JSON
Private nJson As Long
Private bJsonOk As Boolean

' ---------------------------------------------------------------------------
' Point d'entrée (bouton, menu)
' ---------------------------------------------------------------------------
Sub CalcMap()
	Dim oDocCalc As Object, oPlage As Object, aData As Variant
	Dim sUrl As String, oDoc As Object, nOk As Long, bDuCache As Boolean, aPalette As Variant
	Dim bInserer As Boolean, oSelection As Object, oForme As Object, bImage As Boolean
	Dim aOptions As Variant, vTexte As Variant, aColonnes As Variant, sLien As String
	Dim oDocLegende As Object
	Dim aProps(1) As New com.sun.star.beans.PropertyValue
	On Error GoTo Erreur

	OuvrirJournal()
	Journal("CalcMap " & VERSION & " — début")
	ChargerTextes()

	oDocCalc = ThisComponent
	If IsNull(oDocCalc) Then
		Avertir(Tr("msg.open_calc"))
		GoTo Fin
	End If
	If Not oDocCalc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
		Avertir(Tr("msg.use_from_calc"))
		GoTo Fin
	End If

	oSelection = PlageSelectionnee(oDocCalc.getCurrentSelection())
	oPlage = PlageDeDonnees(oSelection)
	If IsNull(oPlage) Then
		Avertir(Tr("msg.select_two_columns"))
		GoTo Fin
	End If
	Journal("Plage lue : " & NomPlage(oPlage))
	aData = oPlage.getDataArray()
	aColonnes = NomsColonnes(oPlage)

	sUrl = ChoisirFondDeCarte(oDocCalc, bDuCache, aPalette, bInserer, aOptions, aColonnes)
	If sUrl = "" Then
		Journal("Choix du fond de carte annulé ou impossible.")
		GoTo Fin
	End If
	Journal("Fond de carte : " & ConvertFromURL(sUrl))

	' Une carte du cache s'ouvre comme modèle : l'enregistrer ne modifie pas le cache.
	' Pour une insertion dans Calc, le dessin reste caché puis est fermé.
	aProps(0).Name = "AsTemplate"
	aProps(0).Value = bDuCache
	aProps(1).Name = "Hidden"
	aProps(1).Value = bInserer
	oDoc = StarDesktop.loadComponentFromURL(sUrl, "_blank", 0, aProps())
	If IsNull(oDoc) Then
		Avertir(Tr("msg.cannot_open_map"))
		GoTo Fin
	End If
	If Not oDoc.supportsService("com.sun.star.drawing.DrawingDocument") Then
		Avertir(Tr("msg.not_drawing"))
		GoTo Fin
	End If

	If (aOptions(O_VALEURS) Or aOptions(O_LIBELLES)) And NombreZones(oDoc) > SEUIL_FORMES_CALC Then
		' Des dizaines de milliers de textes : plusieurs minutes, et illisibles
		Journal("Trop de zones (" & NombreZones(oDoc) & ") : ni valeurs ni libellés")
		Informer(Tr("msg.too_many_values", NombreZones(oDoc)))
		aOptions(O_VALEURS) = False
		aOptions(O_LIBELLES) = False
	End If
	' Légende : créée dans Draw ; pour une carte insérée dans Calc, par le
	' recoloriage qui suit l'insertion (les formes copiées ne la reprennent pas)
	If Not bInserer Then oDocLegende = oDoc
	vTexte = ContexteTexte(oDocCalc, oPlage, aData, aOptions, oDocLegende)
	nOk = ColorierDocument(oDoc, aData, aPalette, vTexte)
	If nOk = 0 Then Avertir(Tr("msg.no_match"))
	If bInserer Then
		If nOk > 0 Then
			' Insertion et lien = une seule action dans l'historique d'annulation
			oDocCalc.getUndoManager().enterUndoContext(Tr("undo.insert_map"))
			oForme = InsererDansCalc(oDoc, oDocCalc, oPlage, NomFichier(sUrl), bImage)
			If Not IsNull(oForme) Then
				sLien = CreerLien(oDocCalc, oSelection, oForme, aPalette(P_ID), sUrl, bImage, aOptions)
				If aOptions(O_LEGENDE) And sLien <> "" Then RecalculerLien(oDocCalc, sLien)
			End If
			oDocCalc.getUndoManager().leaveUndoContext()
		End If
		oDoc.close(True)
	End If
	GoTo Fin

Erreur:
	Journal("ERREUR inattendue " & Err & " (ligne " & Erl & ") : " & Error$)
	On Error Resume Next
	Avertir(Tr("msg.partial"))
	Resume Fin
Fin:
	On Error Resume Next
	Journal("Fin.")
	FermerJournal()
End Sub

' ---------------------------------------------------------------------------
' Cœur : colorie les formes d'un document Draw ouvert avec une palette :
' autant de classes d'intervalles égaux que de couleurs, ou dégradé continu
' entre les couleurs de la palette. Renvoie le nombre de zones coloriées.
' vTexte : contexte du coloriage (ContexteTexte) : colonnes, valeurs,
' libellés, légende.
' Ne lève jamais d'erreur.
' ---------------------------------------------------------------------------
Function NombreZones(oDoc As Object) As Long
	Dim j As Long
	On Error Resume Next
	For j = 0 To oDoc.getDrawPages().getCount() - 1
		NombreZones = NombreZones + oDoc.getDrawPages().getByIndex(j).getCount()
	Next j
End Function

Function ColorierDocument(oDoc As Object, aData As Variant, aPalette As Variant, vTexte As Variant) As Long
	Dim aPages() As Object, j As Long
	On Error GoTo Erreur
	ReDim aPages(oDoc.getDrawPages().getCount() - 1) As Object
	For j = 0 To UBound(aPages)
		aPages(j) = oDoc.getDrawPages().getByIndex(j)
	Next j
	ColorierDocument = ColorierConteneurs(aPages, aData, aPalette, vTexte)
	Exit Function
Erreur:
	Journal("ERREUR pendant le coloriage : " & Error$)
	ColorierDocument = 0
End Function

' Colorie les formes de chaque conteneur (page de dessin, groupe…)
Function ColorierConteneurs(aConteneurs As Variant, aData As Variant, aPalette As Variant, _
		vTexte As Variant) As Long
	Dim oMap As Object, dMin As Double, dMax As Double, nValides As Long
	Dim oLibelles As Object, aOptions As Variant, bTexte As Boolean, dEchelle As Double, nZones As Long
	Dim dTailleVal As Double, dTailleLib As Double, sValeur As String, sLibelle As String
	Dim aCouleurs As Variant, nClasses As Integer, oPage As Object, oShape As Object
	Dim bContinu As Boolean, dPosition As Double, nCouleur As Long
	Dim sNom As String, dVal As Double, iClasse As Integer
	Dim i As Long, j As Long, nOk As Long, nSans As Long
	Dim oUtilises As Object, sCle As String, oEnum As Object, nInconnus As Long
	On Error GoTo Erreur
	ColorierConteneurs = 0
	oUtilises = com.sun.star.container.EnumerableMap.create("string", "boolean")
	aOptions = vTexte(T_OPTIONS)

	oMap = com.sun.star.container.EnumerableMap.create("string", "double")
	oLibelles = com.sun.star.container.EnumerableMap.create("string", "string")
	nValides = LireDonnees(aData, oMap, dMin, dMax, oLibelles, aOptions(O_COL_CODE), _
		aOptions(O_COL_VALEUR), aOptions(O_COL_LIBELLE))
	If nValides = 0 Then
		Journal("Aucune donnée exploitable : la carte reste en couleur par défaut.")
	Else
		Journal("Min = " & dMin & " ; max = " & dMax)
	End If

	aCouleurs = aPalette(P_COULEURS)
	nClasses = UBound(aCouleurs) + 1
	bContinu = aPalette(P_CONTINU)
	If bContinu Then
		Journal("Palette : " & aPalette(P_ID) & " (dégradé continu, " & nClasses & " couleurs)")
	Else
		Journal("Palette : " & aPalette(P_ID) & " (" & nClasses & " classes)")
	End If

	' Tailles des textes (0 : non affichés) : celles choisies pour la carte
	' insérée dans Calc, proportionnelles à la largeur de la carte pour garder
	' le même aspect dans Draw. Le texte des zones est toujours géré (un texte
	' désactivé est effacé), sauf au-delà de SEUIL_FORMES_CALC zones.
	For j = 0 To UBound(aConteneurs)
		nZones = nZones + aConteneurs(j).getCount()
	Next j
	dEchelle = LargeurZones(aConteneurs(0)) / LARGEUR_CARTE_CALC
	If nZones > SEUIL_FORMES_CALC Then
		Journal("Trop de zones (" & nZones & ") : ni valeurs ni libellés")
	Else
		bTexte = True
		If aOptions(O_VALEURS) Then dTailleVal = aOptions(O_TAILLE_VALEURS) * dEchelle
		If aOptions(O_LIBELLES) Then dTailleLib = aOptions(O_TAILLE_LIBELLES) * dEchelle
	End If
	Journal("Options : " & TexteOptions(aOptions))

	For j = 0 To UBound(aConteneurs)
		oPage = aConteneurs(j)
		For i = 0 To oPage.getCount() - 1
			oShape = oPage.getByIndex(i)
			sNom = ""
			On Error Resume Next
			sNom = oShape.getName()
			On Error GoTo Erreur
			If sNom <> "" And Not EstLegende(sNom) Then
				sLibelle = ""
				If dTailleLib > 0 Then sLibelle = TrouverLibelle(oLibelles, sNom)
				If TrouverValeur(oMap, sNom, dVal, sCle) Then
					oUtilises.put(sCle, True)
					' Position de la valeur entre le min (0) et le max (1)
					If dMax = dMin Then
						dPosition = 1
					Else
						dPosition = (dVal - dMin) / (dMax - dMin)
					End If
					If bContinu Then
						nCouleur = CouleurDegrade(aCouleurs, dPosition)
					Else
						iClasse = Int(dPosition * nClasses)
						If iClasse > nClasses - 1 Then iClasse = nClasses - 1
						If iClasse < 0 Then iClasse = 0
						nCouleur = aCouleurs(iClasse)
					End If
					sValeur = ""
					If dTailleVal > 0 Then sValeur = TexteValeur(vTexte, dVal)
					AppliquerZone(oShape, nCouleur, Tr("alt.value", dVal), bTexte, sLibelle, dTailleLib, _
						sValeur, dTailleVal, aOptions)
					nOk = nOk + 1
					If JOURNAL_DETAILLE Then Journal("  " & sNom & " = " & dVal & " -> " _
						& IIf(bContinu, "couleur " & Hex(nCouleur), "classe " & (iClasse + 1)))
				Else
					AppliquerZone(oShape, aPalette(P_VIDE), Tr("alt.none"), bTexte, sLibelle, dTailleLib, _
						"", 0, aOptions)
					nSans = nSans + 1
					If JOURNAL_DETAILLE Then Journal("  " & sNom & " : pas de donnée -> couleur par défaut")
				End If
			End If
		Next i
	Next j

	' Codes présents dans les données mais absents de la carte
	oEnum = oMap.createKeyEnumeration(False)
	Do While oEnum.hasMoreElements()
		sCle = oEnum.nextElement()
		If Not oUtilises.containsKey(sCle) Then
			Journal("Code '" & sCle & "' absent du fond de carte -> ignoré")
			nInconnus = nInconnus + 1
		End If
	Loop
	Journal("Zones coloriées : " & nOk & " ; sans donnée (couleur par défaut) : " & nSans _
		& " ; codes inconnus de la carte : " & nInconnus)

	' Légende : retirée puis recréée (couleurs et bornes à jour)
	If Not IsNull(vTexte(T_DOCUMENT)) Then
		SupprimerLegende(aConteneurs(0))
		If aOptions(O_LEGENDE) Then CreerLegende(vTexte, aConteneurs(0), aPalette, dMin, dMax, _
			nValides > 0, nSans > 0, dEchelle)
	End If
	ColorierConteneurs = nOk
	Exit Function
Erreur:
	Journal("ERREUR pendant le coloriage " & Err & " (ligne " & Erl & ") : " & Error$)
	ColorierConteneurs = nOk
End Function

' Couleur du dégradé à la position dPosition (0 à 1), interpolée
' linéairement entre les couleurs successives de la palette
Function CouleurDegrade(aCouleurs As Variant, ByVal dPosition As Double) As Long
	Dim nSegments As Integer, i As Integer, f As Double, c1 As Long, c2 As Long
	If dPosition < 0 Then dPosition = 0
	If dPosition > 1 Then dPosition = 1
	nSegments = UBound(aCouleurs)
	i = Int(dPosition * nSegments)
	If i > nSegments - 1 Then i = nSegments - 1
	f = dPosition * nSegments - i
	c1 = aCouleurs(i)
	c2 = aCouleurs(i + 1)
	CouleurDegrade = RGB(Int(Red(c1) + (Red(c2) - Red(c1)) * f + 0.5), _
		Int(Green(c1) + (Green(c2) - Green(c1)) * f + 0.5), _
		Int(Blue(c1) + (Blue(c2) - Blue(c1)) * f + 0.5))
End Function

' Applique couleur et texte alternatif, et si bTexte le texte de la zone
' (libellé et valeur ; vides : texte effacé) ; une forme récalcitrante est
' ignorée
Sub AppliquerZone(oShape As Object, nCouleur As Long, sTexte As String, bTexte As Boolean, _
		sLibelle As String, dTailleLib As Double, sValeur As String, dTailleVal As Double, aOptions As Variant)
	Dim aTextes(1) As String, aTailles(1) As Double, aPolices(1) As String, aStyles(1) As String
	Dim n As Integer
	On Error GoTo Erreur
	oShape.FillStyle = com.sun.star.drawing.FillStyle.SOLID
	oShape.FillColor = nCouleur
	If ECRIRE_TEXTE_ALT Then oShape.Description = TexteAlternatif(oShape.Description, sTexte)
	If Not bTexte Then Exit Sub
	n = 0
	If sLibelle <> "" Then
		aTextes(n) = sLibelle : aTailles(n) = dTailleLib
		aPolices(n) = aOptions(O_POLICE_LIBELLES) : aStyles(n) = aOptions(O_STYLE_LIBELLES)
		n = n + 1
	End If
	If sValeur <> "" Then
		aTextes(n) = sValeur : aTailles(n) = dTailleVal
		aPolices(n) = aOptions(O_POLICE_VALEURS) : aStyles(n) = aOptions(O_STYLE_VALEURS)
		n = n + 1
	End If
	If n = 0 Then
		If oShape.getString() <> "" Then oShape.setString("")
	Else
		ReDim Preserve aTextes(n - 1) As String
		ReDim Preserve aTailles(n - 1) As Double
		ReDim Preserve aPolices(n - 1) As String
		ReDim Preserve aStyles(n - 1) As String
		EcrireParagraphes(oShape, aTextes, aTailles, aPolices, aStyles, CouleurTexte(nCouleur))
	End If
	Exit Sub
Erreur:
	Journal("  forme ignorée (" & Error$ & ")")
End Sub

' Ajoute la valeur au texte alternatif existant :
' "22 - Côtes-d'Armor" -> "22 - Côtes-d'Armor - Valeur : 666".
' Une valeur ajoutée par une exécution précédente est remplacée.
Function TexteAlternatif(sExistant As String, sValeur As String) As String
	Dim sPrefixe As String, sBase As String, p As Long
	sBase = sExistant
	sPrefixe = Tr("alt.value", "")
	If sPrefixe <> "" Then
		If Left(sBase, Len(sPrefixe)) = sPrefixe Then
			sBase = ""
		Else
			p = InStr(1, sBase, " - " & sPrefixe, 0)
			If p > 0 Then sBase = Left(sBase, p - 1)
		End If
	End If
	If sBase = "" Then
		TexteAlternatif = sValeur
	Else
		TexteAlternatif = sBase & " - " & sValeur
	End If
End Function

' ---------------------------------------------------------------------------
' Contexte d'un coloriage (voir T_*) : le format de nombre de la première
' valeur numérique de la colonne des valeurs, pour les afficher comme dans la
' feuille (décimales, %, milliers…), les options, le document où créer la
' légende (Nothing : pas de légende) et son titre (en-tête de la colonne).
' ---------------------------------------------------------------------------
Function ContexteTexte(oDocCalc As Object, oPlage As Object, aData As Variant, _
		aOptions As Variant, oDocLegende As Object) As Variant
	Dim oFormateur As Object, nCle As Long, i As Long, iVal As Integer, sTitre As String, d As Double
	iVal = aOptions(O_COL_VALEUR)
	On Error GoTo Erreur
	oFormateur = CreateUnoService("com.sun.star.util.NumberFormatter")
	oFormateur.attachNumberFormatsSupplier(oDocCalc)
	nCle = 0
	For i = 0 To UBound(aData)
		If UBound(aData(i)) >= iVal Then
			Select Case VarType(aData(i)(iVal))
				Case 2, 3, 4, 5, 6, 7
					nCle = oPlage.getCellByPosition(iVal, i).NumberFormat
					Exit For
			End Select
		End If
	Next i
	' En-tête : 1re ligne de la colonne des valeurs, si c'est du texte
	If UBound(aData) >= 0 Then
		If UBound(aData(0)) >= iVal Then
			If VarType(aData(0)(iVal)) = 8 Then
				If Not EnNombre(aData(0)(iVal), d) Then sTitre = Trim(aData(0)(iVal))
			End If
		End If
	End If
	ContexteTexte = Array(oFormateur, nCle, aOptions, oDocLegende, sTitre)
	Exit Function
Erreur:
	Journal("Format des valeurs non lu (" & Error$ & ") : format standard")
	ContexteTexte = Array(Nothing, 0, aOptions, oDocLegende, sTitre)
End Function

' Options de texte (voir O_*), tailles ramenées dans les bornes
Function OptionsParDefaut() As Variant
	OptionsParDefaut = Array(False, TAILLE_TEXTE, False, TAILLE_TEXTE, POLICE_TEXTE, "", POLICE_TEXTE, "", _
		0, 1, 2, False, True)
End Function

Function NomOption(k As Integer) As String
	NomOption = Split(NOMS_OPTIONS, "|")(k)
End Function

' Option k lue depuis un texte (lien, préférences)
Function OptionDeTexte(k As Integer, s As String) As Variant
	Dim aDefaut As Variant
	Select Case Mid(TYPES_OPTIONS, k + 1, 1)
		Case "b"
			OptionDeTexte = (s = "1")
		Case "n"
			OptionDeTexte = TailleValide(Val(s))
		Case "c"
			If Val(s) < 0 Or Val(s) >= NB_COLONNES_MAX Or s = "" Then
				aDefaut = OptionsParDefaut()
				OptionDeTexte = aDefaut(k)
			Else
				OptionDeTexte = CInt(Val(s))
			End If
		Case Else
			OptionDeTexte = s
	End Select
End Function

' Option k écrite en texte (sans |, séparateur du lien)
Function TexteDOption(k As Integer, v As Variant) As String
	If Mid(TYPES_OPTIONS, k + 1, 1) = "b" Then
		TexteDOption = IIf(v, "1", "0")
	Else
		TexteDOption = Replace(CStr(v), "|", "")
	End If
End Function

Function TexteValeur(vTexte As Variant, ByVal dVal As Double) As String
	On Error GoTo Erreur
	TexteValeur = vTexte(T_FORMATEUR).convertNumberToString(vTexte(T_FORMAT), dVal)
	Exit Function
Erreur:
	TexteValeur = CStr(dVal)
End Function

' Taille des valeurs ramenée entre TAILLE_TEXTE_MIN et TAILLE_TEXTE_MAX
' (TAILLE_TEXTE si absente ou illisible)
Function TailleValide(v As Variant) As Integer
	Dim d As Double
	On Error GoTo Defaut
	d = CDbl(v)
	If d <= 0 Then GoTo Defaut
	If d < TAILLE_TEXTE_MIN Then d = TAILLE_TEXTE_MIN
	If d > TAILLE_TEXTE_MAX Then d = TAILLE_TEXTE_MAX
	TailleValide = CInt(d)
	Exit Function
Defaut:
	TailleValide = TAILLE_TEXTE
End Function

' Largeur (1/100 mm) d'une page de dessin, d'une forme ou d'une collection
Function LargeurZones(o As Object) As Double
	Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
	On Error GoTo Defaut
	If o.supportsService("com.sun.star.drawing.DrawPage") Then
		LargeurZones = o.Width
		Exit Function
	End If
	EtendueZones(o, x1, y1, x2, y2)
	If x2 > x1 Then LargeurZones = x2 - x1 Else LargeurZones = LARGEUR_CARTE_CALC
	Exit Function
Defaut:
	LargeurZones = LARGEUR_CARTE_CALC
End Function

' Rectangle (1/100 mm) occupé par les zones d'un conteneur, légende exclue
Sub EtendueZones(o As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
	Dim i As Long, oZ As Object, bInit As Boolean, p As Variant, t As Variant
	For i = 0 To o.getCount() - 1
		oZ = o.getByIndex(i)
		If Not EstLegende(oZ.Name) Then
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

' Texte de la zone : un paragraphe par texte (libellé, valeur), chacun avec
' sa taille, sa police et son style, centré ; le bloc est centré dans la
' zone, sans retour à la ligne et sans agrandir la forme
Sub EcrireParagraphes(oShape As Object, aTextes As Variant, aTailles As Variant, aPolices As Variant, _
		aStyles As Variant, nCouleur As Long)
	Dim oEnum As Object, oPara As Object, k As Integer
	oShape.setString(Join(aTextes, Chr(10)))
	oShape.TextAutoGrowHeight = False
	oShape.TextAutoGrowWidth = False
	oShape.TextWordWrap = False
	oShape.TextHorizontalAdjust = com.sun.star.drawing.TextHorizontalAdjust.CENTER
	oShape.TextVerticalAdjust = com.sun.star.drawing.TextVerticalAdjust.CENTER
	oShape.CharColor = nCouleur
	oEnum = oShape.getText().createEnumeration()
	k = 0
	Do While oEnum.hasMoreElements() And k <= UBound(aTailles)
		oPara = oEnum.nextElement()
		oPara.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER
		AppliquerPolice(oPara, aPolices(k), aTailles(k), aStyles(k))
		k = k + 1
	Loop
End Sub

' Police, taille et style (lettres B, I, U) d'un texte ou d'un paragraphe
Sub AppliquerPolice(o As Object, sPolice As String, dTaille As Double, sStyle As String)
	If sPolice <> "" Then o.CharFontName = sPolice
	o.CharHeight = dTaille
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

' Style (lettres B, I, U) d'un texte ou d'un paragraphe
Function StyleDe(o As Object) As String
	Dim sStyle As String
	If o.CharWeight > com.sun.star.awt.FontWeight.NORMAL Then sStyle = sStyle & "B"
	If o.CharPosture <> com.sun.star.awt.FontSlant.NONE Then sStyle = sStyle & "I"
	If o.CharUnderline <> com.sun.star.awt.FontUnderline.NONE Then sStyle = sStyle & "U"
	StyleDe = sStyle
End Function

' Recopie le texte d'une zone (paragraphes, polices, styles et tailles,
' multipliées par dEchelle)
Sub CopierTexte(oSource As Object, oCopie As Object, dEchelle As Double)
	Dim oEnum As Object, oPara As Object, n As Integer
	Dim aTextes() As String, aTailles() As Double, aPolices() As String, aStyles() As String
	oEnum = oSource.getText().createEnumeration()
	Do While oEnum.hasMoreElements()
		oPara = oEnum.nextElement()
		ReDim Preserve aTextes(n) As String
		ReDim Preserve aTailles(n) As Double
		ReDim Preserve aPolices(n) As String
		ReDim Preserve aStyles(n) As String
		aTextes(n) = oPara.getString()
		aTailles(n) = oPara.CharHeight * dEchelle
		aPolices(n) = oPara.CharFontName
		aStyles(n) = StyleDe(oPara)
		n = n + 1
	Loop
	If n > 0 Then EcrireParagraphes(oCopie, aTextes, aTailles, aPolices, aStyles, oSource.CharColor)
End Sub

' ---------------------------------------------------------------------------
' Légende : une case par classe (5 nuances pour un dégradé continu), avec ses
' bornes si demandé, et « Sans donnée » ; à droite des zones, dans le même
' conteneur (groupe de la carte insérée, ou page de dessin). Toutes ses formes
' sont nommées PREFIXE_LEGENDE.
' ---------------------------------------------------------------------------
Function EstLegende(sNom As String) As Boolean
	EstLegende = (Left(sNom, Len(PREFIXE_LEGENDE)) = PREFIXE_LEGENDE)
End Function

Function AccepteLegende(oConteneur As Object) As Boolean
	On Error GoTo Erreur
	AccepteLegende = oConteneur.supportsService("com.sun.star.drawing.GroupShape") _
		Or oConteneur.supportsService("com.sun.star.drawing.DrawPage")
	Exit Function
Erreur:
	AccepteLegende = False
End Function

Sub SupprimerLegende(oConteneur As Object)
	Dim i As Long
	On Error GoTo Erreur
	If Not AccepteLegende(oConteneur) Then Exit Sub
	For i = oConteneur.getCount() - 1 To 0 Step -1
		If EstLegende(oConteneur.getByIndex(i).Name) Then oConteneur.remove(oConteneur.getByIndex(i))
	Next i
	Exit Sub
Erreur:
	Journal("ERREUR à la suppression de la légende : " & Error$)
End Sub

Sub CreerLegende(vTexte As Variant, oConteneur As Object, aPalette As Variant, dMin As Double, _
		dMax As Double, bDonnees As Boolean, bSans As Boolean, dEchelle As Double)
	Dim oDoc As Object, aOptions As Variant, aCouleurs As Variant, bContinu As Boolean
	Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x As Double, y As Double, xMax As Double
	Dim dTaille As Double, dCase As Double, dPas As Double, k As Integer, n As Integer
	Dim nCouleur As Long, dPos As Double, sTexte As String, bValeurs As Boolean
	On Error GoTo Erreur
	If Not AccepteLegende(oConteneur) Then
		Journal("Légende impossible pour cette carte")
		Exit Sub
	End If
	oDoc = vTexte(T_DOCUMENT)
	aOptions = vTexte(T_OPTIONS)
	bValeurs = aOptions(O_LEGENDE_VALEURS)
	aCouleurs = aPalette(P_COULEURS)
	bContinu = aPalette(P_CONTINU)
	EtendueZones(oConteneur, x1, y1, x2, y2)
	' Texte à la police et à la taille des valeurs (sans leur style, titre en
	' gras) ; case un peu plus haute
	' que le texte (1 pt = 35,28 centièmes de mm)
	dTaille = aOptions(O_TAILLE_VALEURS) * dEchelle
	dCase = dTaille * 35.28 * 1.3
	dPas = dCase * 1.3
	x = x2 + dCase
	y = y1
	xMax = x
	If vTexte(T_TITRE) <> "" And bValeurs Then
		xMax = TexteLegende(oDoc, oConteneur, vTexte(T_TITRE), x, y, dCase, dTaille, _
			aOptions(O_POLICE_VALEURS), "B", xMax)
		y = y + dPas
	End If
	If bDonnees Then
		If bContinu Then n = 5 Else n = UBound(aCouleurs) + 1
		For k = 0 To n - 1
			If bContinu Then
				dPos = k / (n - 1)
				nCouleur = CouleurDegrade(aCouleurs, dPos)
				sTexte = TexteValeur(vTexte, dMin + dPos * (dMax - dMin))
			Else
				nCouleur = aCouleurs(k)
				sTexte = TexteValeur(vTexte, dMin + (dMax - dMin) * k / n) & " – " _
					& TexteValeur(vTexte, dMin + (dMax - dMin) * (k + 1) / n)
			End If
			CaseLegende(oDoc, oConteneur, nCouleur, x, y, dCase * 1.5, dCase)
			If bValeurs Then xMax = TexteLegende(oDoc, oConteneur, sTexte, x + dCase * 1.9, y, dCase, _
				dTaille, aOptions(O_POLICE_VALEURS), "", xMax)
			y = y + dPas
		Next k
	End If
	If bSans Then
		CaseLegende(oDoc, oConteneur, aPalette(P_VIDE), x, y, dCase * 1.5, dCase)
		If bValeurs Then xMax = TexteLegende(oDoc, oConteneur, Tr("legend.no_data"), x + dCase * 1.9, y, _
			dCase, dTaille, aOptions(O_POLICE_VALEURS), "", xMax)
	End If
	' Page de dessin : élargie si la légende dépasse
	If oConteneur.supportsService("com.sun.star.drawing.DrawPage") Then
		If xMax + dCase > oConteneur.Width Then oConteneur.Width = xMax + dCase
	End If
	Exit Sub
Erreur:
	Journal("ERREUR à la création de la légende " & Err & " (ligne " & Erl & ") : " & Error$)
End Sub

Sub CaseLegende(oDoc As Object, oConteneur As Object, nCouleur As Long, x As Double, y As Double, _
		dLargeur As Double, dHauteur As Double)
	Dim oCase As Object
	Dim oPos As New com.sun.star.awt.Point, oTaille As New com.sun.star.awt.Size
	oCase = oDoc.createInstance("com.sun.star.drawing.RectangleShape")
	oConteneur.add(oCase)
	oCase.Name = PREFIXE_LEGENDE
	oPos.X = x : oPos.Y = y
	oTaille.Width = dLargeur : oTaille.Height = dHauteur
	oCase.setPosition(oPos)
	oCase.setSize(oTaille)
	oCase.FillStyle = com.sun.star.drawing.FillStyle.SOLID
	oCase.FillColor = nCouleur
	oCase.LineColor = RGB(96, 96, 96)
End Sub

' Texte de légende, centré verticalement sur une case de hauteur dHauteur ;
' renvoie le bord droit le plus à droite de la légende
Function TexteLegende(oDoc As Object, oConteneur As Object, sTexte As String, x As Double, y As Double, _
		dHauteur As Double, dTaille As Double, sPolice As String, sStyle As String, xMax As Double) As Double
	Dim oTexte As Object
	Dim oPos As New com.sun.star.awt.Point, oTaille As New com.sun.star.awt.Size
	oTexte = oDoc.createInstance("com.sun.star.drawing.TextShape")
	oConteneur.add(oTexte)
	oTexte.Name = PREFIXE_LEGENDE
	oTexte.TextAutoGrowWidth = True
	oTexte.TextAutoGrowHeight = False
	oTexte.TextWordWrap = False
	oTexte.TextVerticalAdjust = com.sun.star.drawing.TextVerticalAdjust.CENTER
	' Ancré à gauche : le texte s'élargit vers la droite (sinon des deux côtés)
	oTexte.TextHorizontalAdjust = com.sun.star.drawing.TextHorizontalAdjust.LEFT
	oTexte.TextLeftDistance = 0
	oTexte.TextRightDistance = 0
	oTexte.setString(sTexte)
	oTexte.CharColor = RGB(0, 0, 0)
	AppliquerPolice(oTexte, sPolice, dTaille, sStyle)
	oPos.X = x : oPos.Y = y
	oTaille.Width = dHauteur : oTaille.Height = dHauteur
	oTexte.setSize(oTaille)
	oTexte.setPosition(oPos)
	TexteLegende = xMax
	If x + oTexte.Size.Width > xMax Then TexteLegende = x + oTexte.Size.Width
End Function

' Noir sur fond clair, blanc sur fond foncé
Function CouleurTexte(nFond As Long) As Long
	If 0.299 * Red(nFond) + 0.587 * Green(nFond) + 0.114 * Blue(nFond) < 140 Then
		CouleurTexte = RGB(255, 255, 255)
	Else
		CouleurTexte = RGB(0, 0, 0)
	End If
End Function

' ---------------------------------------------------------------------------
' Lecture des données : remplit oMap (code -> valeur) et oLibelles (code ->
' libellé) depuis les colonnes iCode, iVal, iLib. Renvoie le nombre de
' lignes retenues. Chaque ligne écartée est tracée dans le journal.
' ---------------------------------------------------------------------------
Function LireDonnees(aData As Variant, oMap As Object, dMin As Double, dMax As Double, _
		oLibelles As Object, iCode As Integer, iVal As Integer, iLib As Integer) As Long
	Dim i As Long, sCode As String, vVal As Variant, dVal As Double, vLib As Variant
	Dim n As Long, nLigne As Long, bInit As Boolean
	On Error GoTo Erreur
	LireDonnees = 0
	If Not IsArray(aData) Then Exit Function

	For i = LBound(aData) To UBound(aData)
		nLigne = i - LBound(aData) + 1
		sCode = "" : vVal = Empty
		If IsArray(aData(i)) Then
			If UBound(aData(i)) >= iCode Then sCode = NormaliserCode(aData(i)(iCode))
			If UBound(aData(i)) >= iVal Then vVal = aData(i)(iVal)
			' Libellé, retenu même sans valeur
			If sCode <> "" And UBound(aData(i)) >= iLib Then
				vLib = aData(i)(iLib)
				If Not EstVide(vLib) Then oLibelles.put(sCode, Trim(CStr(vLib)))
			End If
		End If

		If sCode = "" Then
			If Not EstVide(vVal) Then Journal("Ligne " & nLigne & " ignorée : code vide")
		ElseIf EstVide(vVal) Then
			Journal("Ligne " & nLigne & " (" & sCode & ") : valeur vide -> couleur par défaut")
		ElseIf Not EnNombre(vVal, dVal) Then
			Journal("Ligne " & nLigne & " (" & sCode & ") : valeur non numérique '" & CStr(vVal) _
				& "' -> ignorée (en-tête ?)")
		Else
			If oMap.containsKey(sCode) Then
				Journal("Ligne " & nLigne & " (" & sCode & ") : code en double, la dernière valeur est retenue")
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
			LireDonnees = n
		End If
	Next i
	Journal("Données retenues : " & n & " code(s)")
	Exit Function
Erreur:
	Journal("ERREUR à la lecture des données, ligne " & nLigne & " : " & Error$)
	LireDonnees = n
End Function

' Code -> texte normalisé ("" si inexploitable)
Function NormaliserCode(v As Variant) As String
	Dim s As String
	On Error GoTo Erreur
	NormaliserCode = ""
	Select Case VarType(v)
		Case 2, 3, 4, 5, 6, 7           ' nombres (ex. 75, 1001)
			If v >= 0 And v = Int(v) Then NormaliserCode = Format(v, "0")
		Case 8                           ' texte (ex. "01", "2A")
			s = Trim(Replace(v, Chr(160), " "))
			NormaliserCode = UCase(s)
	End Select
	Exit Function
Erreur:
	NormaliserCode = ""
End Function

Function EstVide(v As Variant) As Boolean
	On Error GoTo Erreur
	EstVide = True
	If IsEmpty(v) Or IsNull(v) Then Exit Function
	If VarType(v) = 8 Then
		If Trim(Replace(v, Chr(160), " ")) = "" Then Exit Function
	End If
	EstVide = False
	Exit Function
Erreur:
	EstVide = True
End Function

' Valeur -> nombre. Accepte les nombres et les textes du type "1 234,5",
' "-12.3", "+7". Renvoie False pour tout le reste (en-têtes, "n.d.", erreurs).
Function EnNombre(v As Variant, dVal As Double) As Boolean
	Dim s As String, c As String, k As Integer, nChiffres As Integer, nSep As Integer
	On Error GoTo Erreur
	EnNombre = False
	Select Case VarType(v)
		Case 2, 3, 4, 5, 6, 7
			dVal = v
			EnNombre = True
			Exit Function
		Case 8
			s = Replace(v, Chr(160), "")
			s = Replace(s, Chr(8239), "")    ' espace fine insécable
			s = Replace(s, " ", "")
			s = Replace(s, ",", ".")
			If s = "" Then Exit Function
			For k = 1 To Len(s)
				c = Mid(s, k, 1)
				If c >= "0" And c <= "9" Then
					nChiffres = nChiffres + 1
				ElseIf c = "." Then
					nSep = nSep + 1
					If nSep > 1 Then Exit Function
				ElseIf (c = "-" Or c = "+") And k = 1 Then
					' signe accepté en tête
				Else
					Exit Function
				End If
			Next k
			If nChiffres = 0 Then Exit Function
			dVal = Val(s)
			EnNombre = True
	End Select
	Exit Function
Erreur:
	EnNombre = False
End Function

' Cherche la valeur d'un code ; tolère les zéros initiaux perdus par Calc
' (ex. département "01" saisi comme le nombre 1, commune "01001" comme 1001)
Function TrouverValeur(oMap As Object, sNom As String, dVal As Double, sCleTrouvee As String) As Boolean
	On Error GoTo Erreur
	TrouverValeur = False
	sCleTrouvee = CleTrouvee(oMap, sNom)
	If sCleTrouvee = "" Then Exit Function
	dVal = oMap.get(sCleTrouvee)
	TrouverValeur = True
	Exit Function
Erreur:
	TrouverValeur = False
End Function

' Libellé d'une zone ("" si absent)
Function TrouverLibelle(oLibelles As Object, sNom As String) As String
	Dim sCle As String
	On Error GoTo Erreur
	sCle = CleTrouvee(oLibelles, sNom)
	If sCle <> "" Then TrouverLibelle = oLibelles.get(sCle)
	Exit Function
Erreur:
	TrouverLibelle = ""
End Function

' Clé de oMap correspondant au nom d'une zone, en tolérant les zéros
' initiaux perdus ("" si absente)
Function CleTrouvee(oMap As Object, sNom As String) As String
	Dim sCle As String
	CleTrouvee = ""
	sCle = UCase(Trim(sNom))
	Do While sCle <> ""
		If oMap.containsKey(sCle) Then
			CleTrouvee = sCle
			Exit Function
		End If
		If Left(sCle, 1) <> "0" Then Exit Do
		sCle = Mid(sCle, 2)
	Loop
End Function

' ---------------------------------------------------------------------------
' Sélection : PlageSelectionnee renvoie les 2 premières colonnes de la
' sélection telles quelles (colonnes entières comprises) ; PlageDeDonnees
' les limite à la zone utilisée de la feuille. Null si inutilisable.
' ---------------------------------------------------------------------------
Function PlageSelectionnee(oSel As Object) As Object
	Dim oPlage As Object, a As Variant, nFin As Long
	On Error GoTo Erreur
	PlageSelectionnee = Nothing
	If IsNull(oSel) Then Exit Function
	If oSel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
		If oSel.getCount() = 0 Then Exit Function
		oPlage = oSel.getByIndex(0)
		Journal("Sélection multiple : seule la première plage est utilisée.")
	ElseIf oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then
		oPlage = oSel
	Else
		Exit Function
	End If
	a = oPlage.getRangeAddress()
	If a.EndColumn - a.StartColumn < 1 Then Exit Function
	' Au moins 3 colonnes (code, valeur, libellé : la 3e est prise même hors de
	' la sélection), au plus NB_COLONNES_MAX
	nFin = a.EndColumn
	If nFin < a.StartColumn + 2 Then nFin = a.StartColumn + 2
	If nFin > a.StartColumn + NB_COLONNES_MAX - 1 Then nFin = a.StartColumn + NB_COLONNES_MAX - 1
	PlageSelectionnee = oPlage.getSpreadsheet().getCellRangeByPosition(a.StartColumn, a.StartRow, _
		nFin, a.EndRow)
	Exit Function
Erreur:
	Journal("ERREUR sur la sélection : " & Error$)
	PlageSelectionnee = Nothing
End Function

Function PlageDeDonnees(oPlage As Object) As Object
	Dim oFeuille As Object, oCurseur As Object
	Dim a As Variant, nFin As Long
	On Error GoTo Erreur
	PlageDeDonnees = Nothing
	If IsNull(oPlage) Then Exit Function
	a = oPlage.getRangeAddress()

	oFeuille = oPlage.getSpreadsheet()
	oCurseur = oFeuille.createCursor()
	oCurseur.gotoEndOfUsedArea(False)
	nFin = oCurseur.getRangeAddress().EndRow
	If a.EndRow < nFin Then nFin = a.EndRow
	If a.StartRow > nFin Then Exit Function

	If a.EndColumn < a.StartColumn + 2 Then a.EndColumn = a.StartColumn + 2   ' liens des versions 0.1 à 0.3
	PlageDeDonnees = oFeuille.getCellRangeByPosition(a.StartColumn, a.StartRow, a.EndColumn, nFin)
	Exit Function
Erreur:
	Journal("ERREUR sur la sélection : " & Error$)
	PlageDeDonnees = Nothing
End Function

' Noms des colonnes d'une plage pour le dialogue : lettre, et en-tête (1re
' ligne) s'il s'agit d'un texte, ex. « B — Valeur »
Function NomsColonnes(oPlage As Object) As Variant
	Dim aNoms() As String, c As Long, n As Long, oCellule As Object, sEntete As String
	On Error GoTo Erreur
	n = oPlage.getColumns().getCount()
	ReDim aNoms(n - 1) As String
	For c = 0 To n - 1
		aNoms(c) = oPlage.getColumns().getByIndex(c).getName()
		oCellule = oPlage.getCellByPosition(c, 0)
		If oCellule.getType() = com.sun.star.table.CellContentType.TEXT Then
			sEntete = Trim(oCellule.getString())
			If Len(sEntete) > 30 Then sEntete = Left(sEntete, 29) & "…"
			If sEntete <> "" Then aNoms(c) = aNoms(c) & " — " & sEntete
		End If
	Next c
	NomsColonnes = aNoms
	Exit Function
Erreur:
	Journal("ERREUR à la lecture des en-têtes : " & Error$)
	NomsColonnes = Array("1", "2", "3")
End Function

Function NomPlage(oPlage As Object) As String
	On Error GoTo Erreur
	NomPlage = oPlage.AbsoluteName
	Exit Function
Erreur:
	NomPlage = "?"
End Function

Function ChoisirFichierOdg() As String
	Dim oPicker As Object
	On Error GoTo Erreur
	ChoisirFichierOdg = ""
	oPicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
	oPicker.appendFilter(Tr("picker.filter"), "*.odg")
	oPicker.setTitle(Tr("picker.title"))
	If oPicker.execute() = 1 Then ChoisirFichierOdg = oPicker.getFiles()(0)
	Exit Function
Erreur:
	Journal("ERREUR dans le choix du fichier : " & Error$)
	ChoisirFichierOdg = ""
End Function

' ---------------------------------------------------------------------------
' Choix du fond de carte (catalogue france-odg) et de la palette. Renvoie
' l'URL du fichier à ouvrir ("" si annulé) ; bDuCache indique une carte du
' cache ; aPalette reçoit la palette choisie (celle par défaut sinon) ;
' bInserer indique si la carte doit être insérée dans la feuille Calc,
' aOptions reçoit les options (colonnes, valeurs, libellés, légende) ;
' aColonnes : noms des colonnes de la sélection.
' ---------------------------------------------------------------------------
Function ChoisirFondDeCarte(oDocCalc As Object, bDuCache As Boolean, aPalette As Variant, _
		bInserer As Boolean, aOptions As Variant, aColonnes As Variant) As String
	Dim sDossier As String, aCartes As Variant, bEnLigne As Boolean, aPalettes As Variant
	Dim sEtat As String, iChoix As Integer, iPalette As Integer, nRes As Integer
	Dim sPrefCarte As String, sPrefPalette As String
	On Error GoTo Erreur
	ChoisirFondDeCarte = ""
	bDuCache = False
	aPalettes = ChargerPalettes()
	LirePreferences(sPrefCarte, sPrefPalette, bInserer, aOptions)
	aPalette = aPalettes(IndicePalette(aPalettes, sPrefPalette))
	sDossier = DossierCartes()
	If sDossier = "" Then
		ChoisirFondDeCarte = ChoisirFichierOdg()
		Exit Function
	End If
	aCartes = ChargerCatalogue(sDossier, bEnLigne)

	If bEnLigne Then
		sEtat = Tr("state.online", ConvertFromURL(sDossier))
	ElseIf UBound(aCartes) >= 0 Then
		sEtat = Tr("state.offline")
	Else
		sEtat = Tr("state.offline_empty")
	End If

	nRes = DialogueCartes(aCartes, aPalettes, sEtat, sPrefCarte, sPrefPalette, _
		iChoix, iPalette, bInserer, aOptions, aColonnes, FenetreParente(oDocCalc), False)
	Journal("Dialogue : résultat " & nRes & ", carte n° " & iChoix & ", palette n° " & iPalette _
		& ", insertion dans Calc : " & bInserer & ", " & TexteOptions(aOptions))
	If (nRes = 1 Or nRes = 3) And iPalette >= 0 Then
		aPalette = aPalettes(iPalette)
		If nRes = 1 And iChoix >= 0 Then
			sPrefCarte = aCartes(iChoix)(C_ID)
			If sPrefCarte = "" Then sPrefCarte = aCartes(iChoix)(C_FICHIER)
		End If
		EcrirePreferences(sPrefCarte, aPalette(P_ID), bInserer, aOptions)
	End If
	If nRes = 1 And iChoix >= 0 Then
		Journal("Carte choisie : " & aCartes(iChoix)(C_TITRE))
		ChoisirFondDeCarte = CarteLocale(aCartes(iChoix), sDossier, bEnLigne, oDocCalc)
		bDuCache = (ChoisirFondDeCarte <> "")
	ElseIf nRes = 3 Or nRes < 0 Then
		' 3 : « Autre fichier… » ; négatif : dialogue non affiché
		ChoisirFondDeCarte = ChoisirFichierOdg()
	End If
	Exit Function
Erreur:
	Journal("ERREUR dans le choix de la carte " & Err & " (ligne " & Erl & ") : " & Error$)
	ChoisirFondDeCarte = ""
End Function

' Options de texte, pour le journal
Function TexteOptions(aOptions As Variant) As String
	Dim k As Integer, s As String
	For k = 0 To NB_OPTIONS - 1
		s = s & IIf(k > 0, ", ", "") & NomOption(k) & "=" & TexteDOption(k, aOptions(k))
	Next k
	TexteOptions = s
End Function

' Derniers choix (preferences.properties du cache) ; insertion dans Calc
' activée par défaut. Les colonnes, propres à chaque tableau, ne sont pas
' mémorisées (1re, 2e et 3e par défaut).
Sub LirePreferences(sCarte As String, sPalette As String, bInserer As Boolean, aOptions As Variant)
	Dim oPrefs As Object, sDossier As String, k As Integer
	On Error GoTo Erreur
	sCarte = ""
	sPalette = ""
	bInserer = True
	aOptions = OptionsParDefaut()
	sDossier = DossierCartes()
	If sDossier = "" Then Exit Sub
	oPrefs = com.sun.star.container.EnumerableMap.create("string", "string")
	LireProprietes(sDossier & NOM_PREFERENCES, oPrefs)
	If oPrefs.containsKey("carte") Then sCarte = oPrefs.get("carte")
	If oPrefs.containsKey("palette") Then sPalette = oPrefs.get("palette")
	If oPrefs.containsKey("inserer") Then bInserer = (oPrefs.get("inserer") <> "0")
	For k = 0 To NB_OPTIONS - 1
		If Mid(TYPES_OPTIONS, k + 1, 1) <> "c" And oPrefs.containsKey(NomOption(k)) Then
			aOptions(k) = OptionDeTexte(k, oPrefs.get(NomOption(k)))
		End If
	Next k
	Exit Sub
Erreur:
	Journal("Préférences illisibles : " & Error$)
End Sub

Sub EcrirePreferences(sCarte As String, sPalette As String, bInserer As Boolean, aOptions As Variant)
	Dim sDossier As String, sTexte As String, k As Integer
	On Error GoTo Erreur
	sDossier = DossierCartes()
	If sDossier = "" Then Exit Sub
	sTexte = "# CalcMap : derniers choix" & Chr(10) & "carte=" & sCarte & Chr(10) _
		& "palette=" & sPalette & Chr(10) & "inserer=" & IIf(bInserer, "1", "0") & Chr(10)
	For k = 0 To NB_OPTIONS - 1
		If Mid(TYPES_OPTIONS, k + 1, 1) <> "c" Then sTexte = sTexte & NomOption(k) & "=" & TexteDOption(k, aOptions(k)) & Chr(10)
	Next k
	EcrireTexte(sDossier & NOM_PREFERENCES, sTexte)
	Exit Sub
Erreur:
	Journal("Préférences non enregistrées : " & Error$)
End Sub

' Dossier du cache ($(user)/calcmap/), créé si besoin. "" si impossible.
' Au premier lancement, le cache de l'ancienne extension ColorierCarte
' (cartes téléchargées, palettes, préférences) est recopié.
Function DossierCartes() As String
	Dim oSFA As Object, sUrl As String, sAncien As String
	On Error GoTo Erreur
	sUrl = CreateUnoService("com.sun.star.util.PathSubstitution") _
		.substituteVariables("$(user)/" & DOSSIER_CARTES, True)
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If Not oSFA.exists(sUrl) Then
		sAncien = CreateUnoService("com.sun.star.util.PathSubstitution") _
			.substituteVariables("$(user)/" & ANCIEN_DOSSIER, True)
		If oSFA.exists(sAncien) Then
			oSFA.copy(sAncien, sUrl)
			Journal("Cache de ColorierCarte repris : " & ConvertFromURL(sAncien))
		Else
			oSFA.createFolder(sUrl)
		End If
	End If
	DossierCartes = sUrl & "/"
	Exit Function
Erreur:
	Journal("ERREUR : dossier des cartes inaccessible (" & Error$ & ")")
	DossierCartes = ""
End Function

' Liste des cartes : catalogue en ligne (copié en cache) ou, sans réseau,
' cartes présentes dans le cache.
Function ChargerCatalogue(sDossier As String, bEnLigne As Boolean) As Variant
	Dim sCat As String, sEmpreinte As String, aCartes As Variant
	On Error GoTo Erreur
	bEnLigne = False
	aCartes = Array()
	sCat = sDossier & NOM_CATALOGUE
	If Telecharger(URL_CATALOGUE, sCat & ".part", sEmpreinte, Nothing, 0) Then
		aCartes = LireCatalogue(sCat & ".part", sDossier)
		If UBound(aCartes) >= 0 Then
			bEnLigne = True
			RemplacerFichier(sCat & ".part", sCat)
			Journal("Catalogue en ligne : " & (UBound(aCartes) + 1) & " carte(s)")
		Else
			Journal("Catalogue en ligne illisible : le cache est utilisé.")
			SupprimerFichier(sCat & ".part")
		End If
	End If
	If Not bEnLigne Then
		aCartes = CartesDuCache(LireCatalogue(sCat, sDossier), sDossier)
		Journal("Hors ligne : " & (UBound(aCartes) + 1) & " carte(s) en cache")
	End If
	ChargerCatalogue = SansCartesIgnorees(aCartes)
	Exit Function
Erreur:
	Journal("ERREUR au chargement du catalogue : " & Error$)
	ChargerCatalogue = Array()
End Function

' Lit un maps.json local ; ignore les entrées incomplètes ou suspectes
Function LireCatalogue(sUrl As String, sDossier As String) As Variant
	Dim oSFA As Object, oRacine As Variant, aListe As Variant, oCarte As Variant
	Dim aRes() As Variant, n As Long, i As Long, sFichier As String, sSource As String
	On Error GoTo Erreur
	LireCatalogue = Array()
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If Not oSFA.exists(sUrl) Then Exit Function
	oRacine = JsonLire(LireTexte(sUrl))
	aListe = JsonElements(JsonChamp(oRacine, "maps"))
	For i = 0 To UBound(aListe)
		oCarte = aListe(i)
		sFichier = JsonTexte(oCarte, "filename")
		sSource = JsonTexte(oCarte, "url")
		If NomFichierSur(sFichier) And LCase(Left(sSource, 8)) = "https://" Then
			ReDim Preserve aRes(n)
			aRes(n) = NouvelleCarte(JsonTexte(oCarte, "id"), JsonTexteTraduit(oCarte, "title"), _
				JsonTexteTraduit(oCarte, "description"), sSource, sFichier, JsonNombre(oCarte, "size"), _
				LCase(JsonTexte(oCarte, "sha256")), oSFA.exists(sDossier & sFichier))
			n = n + 1
		Else
			Journal("Catalogue : entrée " & (i + 1) & " ignorée (fichier ou URL invalide)")
		End If
	Next i
	If n > 0 Then LireCatalogue = aRes
	Exit Function
Erreur:
	Journal("ERREUR à la lecture du catalogue : " & Error$)
	LireCatalogue = Array()
End Function

' Cartes du catalogue présentes en cache, plus les autres .odg du cache
' Retire les cartes de CARTES_IGNOREES (après CartesDuCache, pour que leur
' fichier en cache ne soit pas proposé comme une carte locale inconnue)
Function SansCartesIgnorees(aCartes As Variant) As Variant
	Dim aRes() As Variant, n As Long, i As Long
	SansCartesIgnorees = Array()
	For i = 0 To UBound(aCartes)
		If aCartes(i)(C_ID) <> "" And InStr(CARTES_IGNOREES, "|" & aCartes(i)(C_ID) & "|") > 0 Then
			Journal("Carte ignorée : " & aCartes(i)(C_ID))
		Else
			ReDim Preserve aRes(n)
			aRes(n) = aCartes(i)
			n = n + 1
		End If
	Next i
	If n > 0 Then SansCartesIgnorees = aRes
End Function

Function CartesDuCache(aCatalogue As Variant, sDossier As String) As Variant
	Dim oSFA As Object, aRes() As Variant, n As Long, i As Long, k As Long
	Dim aFichiers As Variant, aParties As Variant, sNom As String, bConnu As Boolean
	On Error GoTo Erreur
	CartesDuCache = Array()
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	For i = 0 To UBound(aCatalogue)
		If aCatalogue(i)(C_LOCALE) Then
			ReDim Preserve aRes(n)
			aRes(n) = aCatalogue(i)
			n = n + 1
		End If
	Next i
	aFichiers = oSFA.getFolderContents(sDossier, False)
	For i = 0 To UBound(aFichiers)
		aParties = Split(ConvertFromURL(aFichiers(i)), GetPathSeparator())
		sNom = aParties(UBound(aParties))
		If LCase(Right(sNom, 4)) = ".odg" Then
			bConnu = False
			For k = 0 To n - 1
				If aRes(k)(C_FICHIER) = sNom Then bConnu = True
			Next k
			If Not bConnu Then
				ReDim Preserve aRes(n)
				aRes(n) = NouvelleCarte("", sNom, Tr("cache.local_map"), "", _
					sNom, oSFA.getSize(aFichiers(i)), "", True)
				n = n + 1
			End If
		End If
	Next i
	If n > 0 Then CartesDuCache = aRes
	Exit Function
Erreur:
	Journal("ERREUR à la lecture du cache : " & Error$)
	If n > 0 Then CartesDuCache = aRes
End Function

Function NouvelleCarte(sId, sTitre, sDescription, sUrl, sFichier, nTaille, sSha256, bLocale) As Variant
	If sTitre = "" Then sTitre = sFichier
	NouvelleCarte = Array(sId, sTitre, sDescription, sUrl, sFichier, nTaille, sSha256, bLocale)
End Function

' Nom de fichier simple en .odg (pas de chemin : le catalogue vient du réseau)
Function NomFichierSur(s As String) As Boolean
	Dim k As Integer, c As String
	NomFichierSur = False
	If Len(s) < 5 Or LCase(Right(s, 4)) <> ".odg" Or Left(s, 1) = "." Then Exit Function
	For k = 1 To Len(s)
		c = Mid(s, k, 1)
		If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._-", c, 0) = 0 Then Exit Function
	Next k
	NomFichierSur = True
End Function

' Renvoie l'URL de la carte en cache, téléchargée ou mise à jour si son
' empreinte diffère de celle du catalogue. "" si aucune version utilisable.
Function CarteLocale(aCarte As Variant, sDossier As String, bEnLigne As Boolean, oDocCalc As Object) As String
	Dim oSFA As Object, sLocal As String, sEmpreinte As String, bExiste As Boolean
	Dim oIndicateur As Object
	On Error GoTo Erreur
	CarteLocale = ""
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	sLocal = sDossier & aCarte(C_FICHIER)
	bExiste = oSFA.exists(sLocal)

	If Not bEnLigne Or aCarte(C_URL) = "" Then
		If bExiste Then
			CarteLocale = sLocal
		Else
			Avertir(Tr("msg.not_offline", aCarte(C_TITRE)))
		End If
		Exit Function
	End If

	If bExiste Then
		If EstAJour(sLocal, aCarte) Then
			Journal("Carte en cache à jour : pas de téléchargement")
			CarteLocale = sLocal
			Exit Function
		End If
		Journal("Carte en cache différente du catalogue : mise à jour")
	Else
		Journal("Carte absente du cache : téléchargement")
	End If

	oIndicateur = IndicateurEtat(oDocCalc)
	If Telecharger(aCarte(C_URL), sLocal & ".part", sEmpreinte, oIndicateur, aCarte(C_TAILLE)) Then
		If aCarte(C_SHA256) = "" Or sEmpreinte = "" Or sEmpreinte = aCarte(C_SHA256) Then
			RemplacerFichier(sLocal & ".part", sLocal)
			Journal("Carte téléchargée (SHA-256 " & sEmpreinte & ")")
			CarteLocale = sLocal
			Exit Function
		End If
		Journal("Empreinte du téléchargement incorrecte : " & sEmpreinte _
			& " au lieu de " & aCarte(C_SHA256))
		SupprimerFichier(sLocal & ".part")
	End If

	If bExiste Then
		Journal("Téléchargement impossible : l'ancienne version en cache est utilisée")
		CarteLocale = sLocal
	Else
		Avertir(Tr("msg.download_failed", aCarte(C_TITRE)))
	End If
	Exit Function
Erreur:
	Journal("ERREUR sur la carte " & aCarte(C_FICHIER) & " : " & Error$)
	If bExiste Then CarteLocale = sLocal
End Function

' Compare le fichier en cache au catalogue : SHA-256, ou taille si le calcul
' d'empreinte est indisponible
Function EstAJour(sLocal As String, aCarte As Variant) As Boolean
	Dim sEmpreinte As String, bIdentique As Boolean
	On Error GoTo Erreur
	EstAJour = False
	If aCarte(C_SHA256) <> "" Then
		sEmpreinte = Sha256Fichier(sLocal)
		If sEmpreinte <> "" Then
			Journal("SHA-256 local " & sEmpreinte & " ; catalogue " & aCarte(C_SHA256))
			EstAJour = (sEmpreinte = aCarte(C_SHA256))
			Exit Function
		End If
	End If
	If aCarte(C_TAILLE) > 0 Then
		bIdentique = (CreateUnoService("com.sun.star.ucb.SimpleFileAccess").getSize(sLocal) = aCarte(C_TAILLE))
		Journal("Comparaison par la taille : " & IIf(bIdentique, "identique", "différente"))
		EstAJour = bIdentique
	Else
		EstAJour = True
	End If
	Exit Function
Erreur:
	EstAJour = False
End Function

' ---------------------------------------------------------------------------
' Téléchargement et empreintes
' ---------------------------------------------------------------------------

' Copie sSource (https ou fichier) vers sDest en calculant son SHA-256.
' False en cas d'échec (pas de réseau, etc.) ; sDest est alors supprimé.
Function Telecharger(sSource As String, sDest As String, sEmpreinte As String, _
		oIndicateur As Object, ByVal nTaille As Double) As Boolean
	Dim oSFA As Object, oEntree As Object, oSortie As Object, oCtx As Object
	Dim aOctets(), n As Long, nLu As Double
	On Error GoTo Erreur
	Telecharger = False
	sEmpreinte = ""
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sDest) Then oSFA.kill(sDest)
	oEntree = oSFA.openFileRead(sSource)
	oSortie = oSFA.openFileWrite(sDest)
	oCtx = ContexteSha256()
	If Not IsNull(oIndicateur) Then oIndicateur.start(Tr("progress.download"), 100)
	Do
		n = oEntree.readBytes(aOctets, 65536)
		If n > 0 Then
			oSortie.writeBytes(aOctets)
			If Not IsNull(oCtx) Then oCtx.updateDigest(aOctets)
			nLu = nLu + n
			If Not IsNull(oIndicateur) And nTaille > 0 Then
				oIndicateur.setValue(IIf(nLu < nTaille, Int(nLu / nTaille * 100), 100))
			End If
		End If
	Loop While n > 0
	oEntree.closeInput()
	oSortie.closeOutput()
	If Not IsNull(oCtx) Then sEmpreinte = EnHexa(oCtx.finalizeDigestAndDispose())
	If Not IsNull(oIndicateur) Then oIndicateur.end()
	Telecharger = True
	Exit Function
Erreur:
	Journal("Téléchargement impossible de " & sSource & " : " & Error$)
	On Error Resume Next
	If Not IsNull(oIndicateur) Then oIndicateur.end()
	If Not IsNull(oEntree) Then oEntree.closeInput()
	If Not IsNull(oSortie) Then oSortie.closeOutput()
	If oSFA.exists(sDest) Then oSFA.kill(sDest)
	Telecharger = False
End Function

' SHA-256 d'un fichier en hexadécimal minuscule ("" si indisponible)
Function Sha256Fichier(sUrl As String) As String
	Dim oEntree As Object, oCtx As Object, aOctets(), n As Long
	On Error GoTo Erreur
	Sha256Fichier = ""
	oCtx = ContexteSha256()
	If IsNull(oCtx) Then Exit Function
	oEntree = CreateUnoService("com.sun.star.ucb.SimpleFileAccess").openFileRead(sUrl)
	Do
		n = oEntree.readBytes(aOctets, 65536)
		If n > 0 Then oCtx.updateDigest(aOctets)
	Loop While n > 0
	oEntree.closeInput()
	Sha256Fichier = EnHexa(oCtx.finalizeDigestAndDispose())
	Exit Function
Erreur:
	Journal("SHA-256 impossible pour " & ConvertFromURL(sUrl) & " : " & Error$)
	On Error Resume Next
	If Not IsNull(oEntree) Then oEntree.closeInput()
	Sha256Fichier = ""
End Function

' Contexte SHA-256 fourni par LibreOffice (NSS) ; Nothing si indisponible
Function ContexteSha256() As Object
	On Error GoTo Erreur
	ContexteSha256 = CreateUnoService("com.sun.star.xml.crypto.NSSInitializer") _
		.getDigestContext(com.sun.star.xml.crypto.DigestID.SHA256, Array())
	Exit Function
Erreur:
	Journal("Calcul SHA-256 indisponible : " & Error$)
	ContexteSha256 = Nothing
End Function

Function EnHexa(aOctets As Variant) As String
	Dim i As Long, s As String
	For i = LBound(aOctets) To UBound(aOctets)
		s = s & Right("0" & LCase(Hex(aOctets(i) And 255)), 2)
	Next i
	EnHexa = s
End Function

Function LireTexte(sUrl As String) As String
	Dim oTexte As Object, s As String
	On Error GoTo Erreur
	oTexte = CreateUnoService("com.sun.star.io.TextInputStream")
	oTexte.setInputStream(CreateUnoService("com.sun.star.ucb.SimpleFileAccess").openFileRead(sUrl))
	oTexte.setEncoding("UTF-8")
	Do While Not oTexte.isEOF()
		s = s & oTexte.readLine() & Chr(10)
	Loop
	oTexte.closeInput()
	LireTexte = s
	Exit Function
Erreur:
	Journal("ERREUR à la lecture de " & ConvertFromURL(sUrl) & " : " & Error$)
	LireTexte = ""
End Function

Sub RemplacerFichier(sDe As String, sVers As String)
	Dim oSFA As Object
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sVers) Then oSFA.kill(sVers)
	oSFA.move(sDe, sVers)
End Sub

Sub SupprimerFichier(sUrl As String)
	Dim oSFA As Object
	On Error Resume Next
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sUrl) Then oSFA.kill(sUrl)
End Sub

' Fenêtre Calc, parente du dialogue (Nothing si indisponible)
Function FenetreParente(oDocCalc As Object) As Object
	On Error GoTo Erreur
	FenetreParente = oDocCalc.getCurrentController().getFrame().getContainerWindow()
	Exit Function
Erreur:
	FenetreParente = Nothing
End Function

' Barre de progression de la fenêtre Calc (Nothing si indisponible)
Function IndicateurEtat(oDocCalc As Object) As Object
	On Error GoTo Erreur
	IndicateurEtat = oDocCalc.getCurrentController().getFrame().createStatusIndicator()
	Exit Function
Erreur:
	IndicateurEtat = Nothing
End Function

Function TailleLisible(nTaille As Double) As String
	If nTaille >= 1048576 Then
		TailleLisible = Tr("size.mb", Format(nTaille / 1048576, "0.0"))
	ElseIf nTaille > 0 Then
		TailleLisible = Tr("size.kb", Format(nTaille / 1024, "0"))
	Else
		TailleLisible = ""
	End If
End Function

' ---------------------------------------------------------------------------
' Dialogue de choix de la carte, de la palette, des colonnes et des options.
' Renvoie 1 (OK), 0 (annuler) ou 3 (autre fichier) ; iChoix et iPalette
' reçoivent les indices choisis, aOptions les options.
' bConfig : configuration d'une carte déjà insérée (fond de carte, insertion
' et « Autre fichier… » non modifiables).
' ---------------------------------------------------------------------------
Function DialogueCartes(aCartes As Variant, aPalettes As Variant, sEtat As String, _
		sPrefCarte As String, sPrefPalette As String, iChoix As Integer, iPalette As Integer, _
		bInserer As Boolean, aOptions As Variant, aColonnes As Variant, oParent As Object, _
		bConfig As Boolean) As Integer
	Dim oModele As Object, oListe As Object, oListePal As Object
	Dim oEcouteurListe As Object, oEcouteurPal As Object, oEcouteurBouton As Object, oEcouteurCases As Object
	Dim aLibelles() As String, aNomsPal() As String, i As Long, iDefaut As Integer, nFin As Long
	Dim iPalDefaut As Integer, aPolices As Variant, sCase As Variant
	On Error GoTo Erreur
	DialogueCartes = 0
	iChoix = -1
	iPalette = -1
	aCartesDialogue = aCartes
	aPalettesDialogue = aPalettes

	nFin = UBound(aCartes)
	If nFin >= 0 Then ReDim aLibelles(nFin) As String
	iDefaut = -1
	For i = 0 To nFin
		aLibelles(i) = aCartes(i)(C_TITRE)
		If Not aCartes(i)(C_LOCALE) Then
			aLibelles(i) = aLibelles(i) & "  — " & Tr("dlg.to_download", TailleLisible(aCartes(i)(C_TAILLE)))
		End If
		If sPrefCarte <> "" And (aCartes(i)(C_ID) = sPrefCarte Or aCartes(i)(C_FICHIER) = sPrefCarte) Then
			iDefaut = i
		ElseIf iDefaut < 0 And aCartes(i)(C_ID) = CARTE_PAR_DEFAUT Then
			iDefaut = i
		End If
	Next i
	If iDefaut < 0 Then iDefaut = 0

	ReDim aNomsPal(UBound(aPalettes)) As String
	For i = 0 To UBound(aPalettes)
		aNomsPal(i) = aPalettes(i)(P_NOM)
	Next i
	iPalDefaut = IndicePalette(aPalettes, sPrefPalette)
	aPolices = PolicesDisponibles(Array(aOptions(O_POLICE_VALEURS), aOptions(O_POLICE_LIBELLES)))

	oModele = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
	oModele.Title = IIf(bConfig, Tr("dlg.config_title"), Tr("app.title"))
	oModele.Width = 330
	oModele.Height = 282
	AjouterControle(oModele, "FixedText", "lblCarte", 6, 6, 318, 10, Array("Label"), Array(Tr("dlg.map")))
	AjouterControle(oModele, "ListBox", "lstCarte", 6, 17, 318, 14, _
		Array("Dropdown", "LineCount", "StringItemList", "Enabled"), Array(True, 12, aLibelles, Not bConfig))
	AjouterControle(oModele, "FixedText", "lblDescription", 6, 35, 318, 30, Array("MultiLine"), Array(True))
	AjouterControle(oModele, "FixedText", "lblPalette", 6, 70, 238, 10, Array("Label"), Array(Tr("dlg.colors")))
	AjouterControle(oModele, "ListBox", "lstPalette", 6, 81, 110, 14, _
		Array("Dropdown", "LineCount", "StringItemList"), Array(True, 16, aNomsPal))
	For i = 0 To NB_NUANCES_APERCU - 1
		AjouterControle(oModele, "FixedText", "nuance" & i, 122, 81, 10, 14, Array("Border"), Array(2))
	Next i
	' Colonnes de la sélection : code, valeur, libellé
	AjouterListeColonne(oModele, "Code", 6, Tr("dlg.col_code"), aColonnes, aOptions(O_COL_CODE))
	AjouterListeColonne(oModele, "Valeur", 114, Tr("dlg.col_value"), aColonnes, aOptions(O_COL_VALEUR))
	AjouterListeColonne(oModele, "Libelle", 222, Tr("dlg.col_label"), aColonnes, aOptions(O_COL_LIBELLE))
	AjouterControle(oModele, "CheckBox", "chkInserer", 6, 130, 318, 12, _
		Array("Label", "State", "Enabled"), Array(Tr("dlg.insert_in_calc"), IIf(bInserer, 1, 0), Not bConfig))
	AjouterLigneTexte(oModele, "Valeurs", 146, Tr("dlg.show_values"), aOptions(O_VALEURS), _
		aOptions(O_TAILLE_VALEURS), aPolices, aOptions(O_POLICE_VALEURS), aOptions(O_STYLE_VALEURS))
	AjouterLigneTexte(oModele, "Libelles", 180, Tr("dlg.show_labels"), aOptions(O_LIBELLES), _
		aOptions(O_TAILLE_LIBELLES), aPolices, aOptions(O_POLICE_LIBELLES), aOptions(O_STYLE_LIBELLES))
	AjouterControle(oModele, "CheckBox", "chkLegende", 6, 214, 150, 12, _
		Array("Label", "State"), Array(Tr("dlg.show_legend"), IIf(aOptions(O_LEGENDE), 1, 0)))
	AjouterControle(oModele, "CheckBox", "chkLegendeValeurs", 160, 214, 164, 12, _
		Array("Label", "State", "Enabled"), Array(Tr("dlg.legend_values"), IIf(aOptions(O_LEGENDE_VALEURS), 1, 0), _
		aOptions(O_LEGENDE)))
	AjouterControle(oModele, "FixedText", "lblEtat", 6, 232, 318, 24, _
		Array("MultiLine", "Label", "TextColor"), Array(True, sEtat, RGB(96, 96, 96)))
	AjouterControle(oModele, "Button", "btnAutre", 6, 262, 70, 14, _
		Array("Label", "PushButtonType"), Array(Tr("dlg.other_file"), 0))
	AjouterControle(oModele, "Button", "btnOk", 218, 262, 50, 14, _
		Array("Label", "PushButtonType", "DefaultButton", "Enabled"), Array(Tr("dlg.ok"), 1, True, nFin >= 0))
	AjouterControle(oModele, "Button", "btnAnnuler", 274, 262, 50, 14, _
		Array("Label", "PushButtonType"), Array(Tr("dlg.cancel"), 2))
	If nFin >= 0 Then oModele.getByName("lstCarte").SelectedItems = Array(CInt(iDefaut))
	oModele.getByName("lstPalette").SelectedItems = Array(CInt(iPalDefaut))

	oDialogue = CreateUnoService("com.sun.star.awt.UnoControlDialog")
	oDialogue.setModel(oModele)
	oListe = oDialogue.getControl("lstCarte")
	oEcouteurListe = CreateUnoListener("ListeCartes_", "com.sun.star.awt.XItemListener")
	oListe.addItemListener(oEcouteurListe)
	oListePal = oDialogue.getControl("lstPalette")
	oEcouteurPal = CreateUnoListener("ListePalettes_", "com.sun.star.awt.XItemListener")
	oListePal.addItemListener(oEcouteurPal)
	oEcouteurBouton = CreateUnoListener("AutreFichier_", "com.sun.star.awt.XActionListener")
	oDialogue.getControl("btnAutre").addActionListener(oEcouteurBouton)
	oEcouteurCases = CreateUnoListener("CasesTexte_", "com.sun.star.awt.XItemListener")
	For Each sCase In Array("chkValeurs", "chkLibelles", "chkLegende")
		oDialogue.getControl(sCase).addItemListener(oEcouteurCases)
	Next sCase
	' Sans peer créé explicitement, execute() renvoie -1 sans rien afficher
	' (constaté avec LibreOffice 25.2 sous Wayland)
	oDialogue.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), oParent)
	If nFin >= 0 Then AfficherDescription(iDefaut)
	AfficherPalette(iPalDefaut)
	If bConfig Then
		oDialogue.getControl("btnAutre").setVisible(False)
		oDialogue.getControl("lblDescription").getModel().Label = Tr("dlg.config_hint")
	End If

	DialogueCartes = oDialogue.execute()
	iChoix = oListe.getSelectedItemPos()
	iPalette = oListePal.getSelectedItemPos()
	bInserer = (oDialogue.getControl("chkInserer").getState() = 1)
	aOptions(O_COL_CODE) = ColonneChoisie("Code", aOptions(O_COL_CODE))
	aOptions(O_COL_VALEUR) = ColonneChoisie("Valeur", aOptions(O_COL_VALEUR))
	aOptions(O_COL_LIBELLE) = ColonneChoisie("Libelle", aOptions(O_COL_LIBELLE))
	LireLigneTexte("Valeurs", aOptions, O_VALEURS, O_TAILLE_VALEURS, O_POLICE_VALEURS, O_STYLE_VALEURS)
	LireLigneTexte("Libelles", aOptions, O_LIBELLES, O_TAILLE_LIBELLES, O_POLICE_LIBELLES, O_STYLE_LIBELLES)
	aOptions(O_LEGENDE) = (oDialogue.getControl("chkLegende").getState() = 1)
	aOptions(O_LEGENDE_VALEURS) = (oDialogue.getControl("chkLegendeValeurs").getState() = 1)
	For Each sCase In Array("chkValeurs", "chkLibelles", "chkLegende")
		oDialogue.getControl(sCase).removeItemListener(oEcouteurCases)
	Next sCase
	If DialogueCartes = 0 Then Journal("Dialogue de choix de carte annulé")
	oListe.removeItemListener(oEcouteurListe)
	oListePal.removeItemListener(oEcouteurPal)
	oDialogue.getControl("btnAutre").removeActionListener(oEcouteurBouton)
	oDialogue.dispose()
	oDialogue = Nothing
	Exit Function
Erreur:
	Journal("ERREUR dans le dialogue de choix de carte " & Err & " (ligne " & Erl & ") : " & Error$)
	On Error Resume Next
	If Not IsNull(oDialogue) Then oDialogue.dispose()
	oDialogue = Nothing
	DialogueCartes = 3          ' repli : sélecteur de fichier classique
End Function

' Case à cocher, taille (flèches +1 / -1, ou saisie directe), puis sur une
' 2e ligne police et style, grisés tant que la case n'est pas cochée :
' chk<sSuffixe>, numTaille<sSuffixe>, lstPolice<sSuffixe>, chkGras<sSuffixe>…
Sub AjouterLigneTexte(oModele As Object, sSuffixe As String, nY As Long, sLibelle As String, _
		bCoche As Boolean, nTaille As Integer, aPolices As Variant, sPolice As String, sStyle As String)
	Dim iPolice As Integer, i As Integer
	AjouterControle(oModele, "CheckBox", "chk" & sSuffixe, 6, nY, 200, 12, _
		Array("Label", "State"), Array(sLibelle, IIf(bCoche, 1, 0)))
	AjouterControle(oModele, "FixedText", "lblTaille" & sSuffixe, 210, nY + 2, 70, 10, _
		Array("Label", "Align", "Enabled"), Array(Tr("dlg.text_size"), 2, bCoche))
	AjouterControle(oModele, "NumericField", "numTaille" & sSuffixe, 284, nY - 1, 40, 14, _
		Array("Value", "ValueMin", "ValueMax", "ValueStep", "DecimalAccuracy", "Spin", "StrictFormat", _
		"Repeat", "Enabled"), _
		Array(CDbl(TailleValide(nTaille)), CDbl(TAILLE_TEXTE_MIN), CDbl(TAILLE_TEXTE_MAX), 1#, 0, True, _
		True, True, bCoche))
	For i = 0 To UBound(aPolices)
		If aPolices(i) = sPolice Then iPolice = i
	Next i
	AjouterControle(oModele, "FixedText", "lblPolice" & sSuffixe, 18, nY + 18, 36, 10, _
		Array("Label", "Enabled"), Array(Tr("dlg.font"), bCoche))
	AjouterControle(oModele, "ListBox", "lstPolice" & sSuffixe, 56, nY + 15, 120, 14, _
		Array("Dropdown", "LineCount", "StringItemList", "Enabled"), Array(True, 16, aPolices, bCoche))
	oModele.getByName("lstPolice" & sSuffixe).SelectedItems = Array(iPolice)
	AjouterControle(oModele, "CheckBox", "chkGras" & sSuffixe, 182, nY + 16, 44, 12, _
		Array("Label", "State", "Enabled"), Array(Tr("dlg.bold"), IIf(InStr(sStyle, "B") > 0, 1, 0), bCoche))
	AjouterControle(oModele, "CheckBox", "chkItalique" & sSuffixe, 228, nY + 16, 48, 12, _
		Array("Label", "State", "Enabled"), Array(Tr("dlg.italic"), IIf(InStr(sStyle, "I") > 0, 1, 0), bCoche))
	AjouterControle(oModele, "CheckBox", "chkSouligne" & sSuffixe, 278, nY + 16, 46, 12, _
		Array("Label", "State", "Enabled"), Array(Tr("dlg.underline"), IIf(InStr(sStyle, "U") > 0, 1, 0), bCoche))
End Sub

' Liste déroulante d'une colonne (lbl<sSuffixe>, lst<sSuffixe>)
Sub AjouterListeColonne(oModele As Object, sSuffixe As String, nX As Long, sLibelle As String, _
		aColonnes As Variant, iDefaut As Integer)
	Dim i As Integer
	i = iDefaut
	If i > UBound(aColonnes) Then i = UBound(aColonnes)
	If i < 0 Then i = 0
	AjouterControle(oModele, "FixedText", "lbl" & sSuffixe, nX, 100, 102, 10, Array("Label"), Array(sLibelle))
	AjouterControle(oModele, "ListBox", "lst" & sSuffixe, nX, 111, 102, 14, _
		Array("Dropdown", "LineCount", "StringItemList"), Array(True, 12, aColonnes))
	oModele.getByName("lst" & sSuffixe).SelectedItems = Array(i)
End Sub

' Valeurs d'une ligne de texte du dialogue (case, taille, police, style)
Sub LireLigneTexte(sSuffixe As String, aOptions As Variant, kCase As Integer, kTaille As Integer, _
		kPolice As Integer, kStyle As Integer)
	Dim sStyle As String, sPolice As String
	aOptions(kCase) = (oDialogue.getControl("chk" & sSuffixe).getState() = 1)
	aOptions(kTaille) = TailleValide(oDialogue.getControl("numTaille" & sSuffixe).getValue())
	sPolice = oDialogue.getControl("lstPolice" & sSuffixe).getSelectedItem()
	If sPolice <> "" Then aOptions(kPolice) = sPolice
	If oDialogue.getControl("chkGras" & sSuffixe).getState() = 1 Then sStyle = sStyle & "B"
	If oDialogue.getControl("chkItalique" & sSuffixe).getState() = 1 Then sStyle = sStyle & "I"
	If oDialogue.getControl("chkSouligne" & sSuffixe).getState() = 1 Then sStyle = sStyle & "U"
	aOptions(kStyle) = sStyle
End Sub

Function ColonneChoisie(sSuffixe As String, iDefaut As Integer) As Integer
	ColonneChoisie = oDialogue.getControl("lst" & sSuffixe).getSelectedItemPos()
	If ColonneChoisie < 0 Then ColonneChoisie = iDefaut
End Function

' Polices installées (et celles de aAjouts), triées
Function PolicesDisponibles(aAjouts As Variant) As Variant
	Dim oNoms As Object, aDesc As Variant, i As Long, aRes() As String, n As Long, oEnum As Object
	oNoms = com.sun.star.container.EnumerableMap.create("string", "boolean")
	On Error GoTo Suite
	aDesc = CreateUnoService("com.sun.star.awt.Toolkit").createScreenCompatibleDevice(16, 16).getFontDescriptors()
	For i = 0 To UBound(aDesc)
		If aDesc(i).Name <> "" Then oNoms.put(aDesc(i).Name, True)
	Next i
Suite:
	On Error Resume Next
	oNoms.put(POLICE_TEXTE, True)
	For i = 0 To UBound(aAjouts)
		If aAjouts(i) <> "" Then oNoms.put(aAjouts(i), True)
	Next i
	oEnum = oNoms.createKeyEnumeration(False)
	Do While oEnum.hasMoreElements()
		ReDim Preserve aRes(n) As String
		aRes(n) = oEnum.nextElement()
		n = n + 1
	Loop
	TrierTextes(aRes)
	PolicesDisponibles = aRes
End Function

' Tri (Shell) de textes, sans distinction de casse
Sub TrierTextes(a As Variant)
	Dim nEcart As Long, i As Long, j As Long, s As String
	nEcart = (UBound(a) + 1) \ 2
	Do While nEcart > 0
		For i = nEcart To UBound(a)
			s = a(i)
			j = i
			Do While j >= nEcart
				If LCase(a(j - nEcart)) <= LCase(s) Then Exit Do
				a(j) = a(j - nEcart)
				j = j - nEcart
			Loop
			a(j) = s
		Next i
		nEcart = nEcart \ 2
	Loop
End Sub

Sub AjouterControle(oModele As Object, sType As String, sNom As String, nX As Long, nY As Long, _
		nLargeur As Long, nHauteur As Long, aNoms As Variant, aValeurs As Variant)
	Dim oCtrl As Object, k As Integer
	oCtrl = oModele.createInstance("com.sun.star.awt.UnoControl" & sType & "Model")
	oCtrl.PositionX = nX
	oCtrl.PositionY = nY
	oCtrl.Width = nLargeur
	oCtrl.Height = nHauteur
	For k = 0 To UBound(aNoms)
		oCtrl.setPropertyValue(aNoms(k), aValeurs(k))
	Next k
	oModele.insertByName(sNom, oCtrl)
End Sub

Sub AfficherDescription(ByVal i As Long)
	Dim aCarte As Variant, s As String
	On Error Resume Next
	If i < 0 Or i > UBound(aCartesDialogue) Then Exit Sub
	aCarte = aCartesDialogue(i)
	s = aCarte(C_DESCRIPTION) & Chr(10) & aCarte(C_FICHIER)
	If aCarte(C_TAILLE) > 0 Then s = s & " — " & TailleLisible(aCarte(C_TAILLE))
	s = s & " — " & IIf(aCarte(C_LOCALE), Tr("dlg.downloaded"), Tr("dlg.will_download"))
	oDialogue.getControl("lblDescription").getModel().Label = s
End Sub

' Aperçu de la palette : une case par classe, ou pour un dégradé continu
' toutes les cases, jointives et sans bordure
Sub AfficherPalette(ByVal i As Long)
	Dim aCouleurs As Variant, n As Integer, k As Integer, nLargeur As Integer, oCase As Object
	Dim bContinu As Boolean, nCouleur As Long
	On Error Resume Next
	If i < 0 Or i > UBound(aPalettesDialogue) Then Exit Sub
	aCouleurs = aPalettesDialogue(i)(P_COULEURS)
	bContinu = aPalettesDialogue(i)(P_CONTINU)
	n = UBound(aCouleurs) + 1
	If bContinu Or n > NB_NUANCES_APERCU Then n = NB_NUANCES_APERCU
	nLargeur = Int(122 / n)
	If nLargeur > 20 Then nLargeur = 20
	For k = 0 To NB_NUANCES_APERCU - 1
		oCase = oDialogue.getControl("nuance" & k)
		If k < n Then
			If bContinu Then
				nCouleur = CouleurDegrade(aCouleurs, k / (n - 1))
			Else
				nCouleur = aCouleurs(k)
			End If
			oCase.getModel().PositionX = 122 + k * nLargeur
			oCase.getModel().Width = nLargeur
			oCase.getModel().Border = IIf(bContinu, 0, 2)
			oCase.getModel().BackgroundColor = nCouleur
			oCase.setVisible(True)
		Else
			oCase.setVisible(False)
		End If
	Next k
End Sub

' Écouteurs du dialogue
Sub ListeCartes_itemStateChanged(oEvenement As Object)
	On Error Resume Next
	AfficherDescription(oEvenement.Selected)
End Sub

Sub ListeCartes_disposing(oEvenement As Object)
End Sub

Sub ListePalettes_itemStateChanged(oEvenement As Object)
	On Error Resume Next
	AfficherPalette(oEvenement.Selected)
End Sub

Sub ListePalettes_disposing(oEvenement As Object)
End Sub

' Cases « valeurs », « libellés » et « légende » : les réglages qui en
' dépendent ne sont modifiables que si la case est cochée
Sub CasesTexte_itemStateChanged(oEvenement As Object)
	Dim sSuffixe As Variant, sCtrl As Variant, bActif As Boolean
	On Error Resume Next
	For Each sSuffixe In Array("Valeurs", "Libelles")
		bActif = (oDialogue.getControl("chk" & sSuffixe).getState() = 1)
		For Each sCtrl In Array("lblTaille", "numTaille", "lblPolice", "lstPolice", "chkGras", "chkItalique", "chkSouligne")
			oDialogue.getControl(sCtrl & sSuffixe).getModel().Enabled = bActif
		Next sCtrl
	Next sSuffixe
	oDialogue.getControl("chkLegendeValeurs").getModel().Enabled = (oDialogue.getControl("chkLegende").getState() = 1)
End Sub

Sub CasesTexte_disposing(oEvenement As Object)
End Sub

Sub AutreFichier_actionPerformed(oEvenement As Object)
	On Error Resume Next
	oDialogue.endDialog(3)
End Sub

Sub AutreFichier_disposing(oEvenement As Object)
End Sub

' ---------------------------------------------------------------------------
' Insertion de la carte coloriée dans la feuille Calc, à droite des données
' ---------------------------------------------------------------------------
Function InsererDansCalc(oDoc As Object, oDocCalc As Object, oPlage As Object, sNom As String, _
		bImage As Boolean) As Object
	Dim oFeuille As Object, oCellule As Object, oPageDraw As Object, oPageCalc As Object
	Dim oForme As Object, nFormes As Long, a As Variant, oBloc As Object
	Dim oTaille As New com.sun.star.awt.Size
	On Error GoTo Erreur
	InsererDansCalc = Nothing
	bImage = False
	oFeuille = oPlage.getSpreadsheet()
	' À droite du bloc de données contigu à la sélection, après une colonne vide
	oBloc = oFeuille.createCursorByRange(oPlage)
	oBloc.collapseToCurrentRegion()
	a = oBloc.getRangeAddress()
	If a.EndColumn < oPlage.getRangeAddress().EndColumn Then a = oPlage.getRangeAddress()
	oCellule = oFeuille.getCellByPosition(a.EndColumn + 2, oPlage.getRangeAddress().StartRow)
	oPageDraw = oDoc.getDrawPages().getByIndex(0)
	oPageCalc = oFeuille.getDrawPage()
	nFormes = oPageDraw.getCount()

	If nFormes <= SEUIL_FORMES_CALC Then oForme = CopierFormes(oPageDraw, oDocCalc, oPageCalc)
	If IsNull(oForme) Then
		oForme = InsererImage(oPageDraw, oDocCalc, oPageCalc)
		bImage = True
	End If
	If IsNull(oForme) Then
		Avertir(Tr("msg.insert_failed"))
		Exit Function
	End If

	' Largeur fixe, proportions conservées, coin supérieur gauche sur la cellule
	oTaille.Width = LARGEUR_CARTE_CALC
	oTaille.Height = LARGEUR_CARTE_CALC * oForme.Size.Height / oForme.Size.Width
	oForme.setSize(oTaille)
	oForme.setPosition(oCellule.Position)
	oForme.Name = NomFormeLibre(oPageCalc, sNom)
	oForme.Description = sNom
	On Error Resume Next
	oForme.Anchor = oCellule
	oDocCalc.getCurrentController().select(oForme)
	On Error GoTo Erreur
	Journal("Carte insérée dans Calc en " & NomCellule(oCellule) & " (" _
		& IIf(bImage, "image", nFormes & " formes") & ")")
	If bImage Then Informer(Tr("msg.inserted_as_image", nFormes))
	InsererDansCalc = oForme
	Exit Function
Erreur:
	Journal("ERREUR à l'insertion dans Calc " & Err & " (ligne " & Erl & ") : " & Error$)
	Avertir(Tr("msg.insert_failed"))
End Function

' Recrée les zones de la carte dans la feuille et les groupe. Chaque zone
' garde géométrie, couleurs, nom et texte alternatif. Nothing si impossible
' (les formes déjà créées sont alors retirées).
Function CopierFormes(oPageDraw As Object, oDocCalc As Object, oPageCalc As Object) As Object
	Dim oColl As Object, oSource As Object, oCopie As Object, sType As String
	Dim i As Long, nIgnorees As Long
	On Error GoTo Erreur
	CopierFormes = Nothing
	oColl = CreateUnoService("com.sun.star.drawing.ShapeCollection")
	For i = 0 To oPageDraw.getCount() - 1
		oSource = oPageDraw.getByIndex(i)
		sType = oSource.getShapeType()
		If sType = "com.sun.star.drawing.PolyPolygonShape" Or sType = "com.sun.star.drawing.ClosedBezierShape" Then
			oCopie = oDocCalc.createInstance(sType)
			oPageCalc.add(oCopie)
			oColl.add(oCopie)
			If sType = "com.sun.star.drawing.PolyPolygonShape" Then
				oCopie.PolyPolygon = oSource.PolyPolygon
			Else
				oCopie.PolyPolygonBezier = oSource.PolyPolygonBezier
			End If
			oCopie.FillStyle = oSource.FillStyle
			oCopie.FillColor = oSource.FillColor
			oCopie.LineStyle = oSource.LineStyle
			oCopie.LineColor = oSource.LineColor
			oCopie.LineWidth = oSource.LineWidth
			oCopie.Name = oSource.Name
			oCopie.Description = oSource.Description
			If oSource.getString() <> "" Then CopierTexte(oSource, oCopie, LARGEUR_CARTE_CALC / oPageDraw.Width)
		Else
			nIgnorees = nIgnorees + 1
		End If
	Next i
	If nIgnorees > 0 Then Journal(nIgnorees & " forme(s) non géographique(s) non copiée(s)")
	If oColl.getCount() = 0 Then Exit Function
	If oColl.getCount() = 1 Then
		CopierFormes = oColl.getByIndex(0)
	Else
		CopierFormes = oPageCalc.group(oColl)
	End If
	Exit Function
Erreur:
	Journal("ERREUR à la copie des formes : " & Error$ & " ; repli sur une image")
	On Error Resume Next
	For i = 0 To oColl.getCount() - 1
		oPageCalc.remove(oColl.getByIndex(i))
	Next i
	CopierFormes = Nothing
End Function

' Exporte la page de la carte en PNG et l'insère dans la feuille.
' Nothing si impossible.
Function InsererImage(oPageDraw As Object, oDocCalc As Object, oPageCalc As Object) As Object
	Dim oGraphique As Object, oImage As Object
	Dim oTaille As New com.sun.star.awt.Size
	On Error GoTo Erreur
	InsererImage = Nothing
	oGraphique = ImageDePage(oPageDraw)
	If IsNull(oGraphique) Then Exit Function
	oImage = oDocCalc.createInstance("com.sun.star.drawing.GraphicObjectShape")
	oPageCalc.add(oImage)
	oImage.Graphic = oGraphique
	oTaille.Width = oPageDraw.Width
	oTaille.Height = oPageDraw.Height
	oImage.setSize(oTaille)
	InsererImage = oImage
	Exit Function
Erreur:
	Journal("ERREUR à l'insertion de l'image : " & Error$)
	InsererImage = Nothing
End Function

' Page de dessin -> image PNG (Graphic) ; Nothing si impossible
Function ImageDePage(oPageDraw As Object) As Object
	Dim oFiltre As Object, sUrl As String
	Dim aDonnees(1) As New com.sun.star.beans.PropertyValue
	Dim aArgs(2) As New com.sun.star.beans.PropertyValue
	Dim aSource(0) As New com.sun.star.beans.PropertyValue
	On Error GoTo Erreur
	ImageDePage = Nothing
	sUrl = DossierCartes() & NOM_IMAGE
	SupprimerFichier(sUrl)
	aDonnees(0).Name = "PixelWidth"
	aDonnees(0).Value = LARGEUR_IMAGE_PX
	aDonnees(1).Name = "PixelHeight"
	aDonnees(1).Value = CLng(LARGEUR_IMAGE_PX * oPageDraw.Height / oPageDraw.Width)
	aArgs(0).Name = "URL"
	aArgs(0).Value = sUrl
	aArgs(1).Name = "MediaType"
	aArgs(1).Value = "image/png"
	aArgs(2).Name = "FilterData"
	aArgs(2).Value = aDonnees()
	oFiltre = CreateUnoService("com.sun.star.drawing.GraphicExportFilter")
	oFiltre.setSourceDocument(oPageDraw)
	oFiltre.filter(aArgs())

	aSource(0).Name = "URL"
	aSource(0).Value = sUrl
	ImageDePage = CreateUnoService("com.sun.star.graphic.GraphicProvider").queryGraphic(aSource())
	Exit Function
Erreur:
	Journal("ERREUR à l'export de l'image : " & Error$)
	ImageDePage = Nothing
End Function

' Nom de forme non utilisé sur la page : "nom", sinon "nom (2)", "nom (3)"…
Function NomFormeLibre(oPage As Object, sNom As String) As String
	Dim n As Integer, s As String
	s = sNom
	n = 1
	Do While Not IsNull(TrouverForme(oPage, s))
		n = n + 1
		s = sNom & " (" & n & ")"
	Loop
	NomFormeLibre = s
End Function

' Forme de premier niveau portant ce nom ; Nothing si absente
Function TrouverForme(oPage As Object, sNom As String) As Object
	Dim i As Long
	On Error GoTo Erreur
	TrouverForme = Nothing
	For i = 0 To oPage.getCount() - 1
		If oPage.getByIndex(i).Name = sNom Then
			TrouverForme = oPage.getByIndex(i)
			Exit Function
		End If
	Next i
Erreur:
End Function

' ---------------------------------------------------------------------------
' Mise à jour automatique des cartes insérées dans Calc.
' Un lien = une plage nommée CalcMap_N (les 2 colonnes sélectionnées,
' qui suit les insertions de lignes) + une propriété du document de même nom :
' "1|nom de la forme|id de palette|URL du fond de carte|image (0/1)|options"
' (ChampsLien ; options dans l'ordre O_*, séparées par | ; les liens des
' versions précédentes en ont moins : valeurs par défaut).
' ---------------------------------------------------------------------------
Function CreerLien(oDocCalc As Object, oSelection As Object, oForme As Object, sPalette As String, _
		sUrlCarte As String, bImage As Boolean, aOptions As Variant) As String
	Dim oPlages As Object, sNom As String, n As Integer, oProps As Object
	Dim oOrigine As New com.sun.star.table.CellAddress
	On Error GoTo Erreur
	CreerLien = ""
	oPlages = oDocCalc.NamedRanges
	n = 1
	Do While oPlages.hasByName(PREFIXE_LIEN & n)
		n = n + 1
	Loop
	sNom = PREFIXE_LIEN & n
	oPlages.addNewByName(sNom, oSelection.AbsoluteName, oOrigine, 0)
	oProps = oDocCalc.getDocumentProperties().getUserDefinedProperties()
	oProps.addProperty(sNom, com.sun.star.beans.PropertyAttribute.REMOVABLE, _
		ChampsLien(oForme.Name, sPalette, sUrlCarte, bImage, aOptions))
	Journal("Carte liée aux données " & oSelection.AbsoluteName & " (" & sNom & ") : mise à jour automatique")
	SuivreLien(oDocCalc, sNom)
	CreerLien = sNom
	Exit Function
Erreur:
	Journal("ERREUR : la carte ne sera pas mise à jour automatiquement (" & Error$ & ")")
End Function

Function ChampsLien(sForme As String, sPalette As String, sUrlCarte As String, bImage As Boolean, _
		aOptions As Variant) As String
	Dim s As String, k As Integer
	s = "1|" & sForme & "|" & sPalette & "|" & sUrlCarte & "|" & IIf(bImage, "1", "0")
	For k = 0 To NB_OPTIONS - 1
		s = s & "|" & TexteDOption(k, aOptions(k))
	Next k
	ChampsLien = s
End Function

' Options d'un lien (champs absents : valeurs par défaut)
Function OptionsDuLien(aChamps As Variant) As Variant
	Dim aOptions As Variant, k As Integer
	aOptions = OptionsParDefaut()
	For k = 0 To NB_OPTIONS - 1
		If UBound(aChamps) >= 5 + k Then aOptions(k) = OptionDeTexte(k, aChamps(5 + k))
	Next k
	OptionsDuLien = aOptions
End Function

' Écoute les modifications de la plage d'un lien (une seule fois par lien)
Sub SuivreLien(oDocCalc As Object, sNom As String)
	Dim oPlage As Object, oEcouteur As Object, i As Long
	On Error GoTo Erreur
	For i = 0 To nLiens - 1
		If EqualUnoObjects(aLiens(i)(0), oDocCalc) And aLiens(i)(3) = sNom Then Exit Sub
	Next i
	oPlage = oDocCalc.NamedRanges.getByName(sNom).getReferredCells()
	If IsNull(oPlage) Then Exit Sub
	oEcouteur = CreateUnoListener("DonneesCarte_", "com.sun.star.util.XModifyListener")
	oPlage.addModifyListener(oEcouteur)
	ReDim Preserve aLiens(nLiens)
	aLiens(nLiens) = Array(oDocCalc, oPlage, oEcouteur, sNom)
	nLiens = nLiens + 1
	IntercepterMenu(oDocCalc)
	Exit Sub
Erreur:
	Journal("ERREUR au suivi de " & sNom & " : " & Error$)
End Sub

' Événement « Ouvrir le document » (Events.xcu) : reconnecte les cartes liées
' d'un classeur. Silencieux et rapide pour tout autre document.
Sub ReconnecterCartes(Optional oEvenement As Variant)
	Dim oDoc As Object, oProps As Object, aNoms As Variant, i As Long
	On Error GoTo Fin
	If IsMissing(oEvenement) Then
		oDoc = ThisComponent
	Else
		oDoc = oEvenement.Source
	End If
	If IsNull(oDoc) Then Exit Sub
	If Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then Exit Sub
	oProps = oDoc.getDocumentProperties().getUserDefinedProperties()
	aNoms = oProps.getPropertySetInfo().getProperties()
	For i = 0 To UBound(aNoms)
		If Left(aNoms(i).Name, Len(PREFIXE_LIEN)) = PREFIXE_LIEN _
				Or Left(aNoms(i).Name, Len(ANCIEN_PREFIXE_LIEN)) = ANCIEN_PREFIXE_LIEN Then
			If oDoc.NamedRanges.hasByName(aNoms(i).Name) Then SuivreLien(oDoc, aNoms(i).Name)
		End If
	Next i
Fin:
End Sub

' Écouteur : une valeur de la plage a changé
Sub DonneesCarte_modified(oEvenement As Object)
	Dim i As Long
	On Error GoTo Fin
	If bRecalculEnCours Then Exit Sub
	For i = 0 To nLiens - 1
		If EqualUnoObjects(aLiens(i)(1), oEvenement.Source) Then
			bRecalculEnCours = True
			RecalculerLien(aLiens(i)(0), aLiens(i)(3))
			bRecalculEnCours = False
			Exit Sub
		End If
	Next i
Fin:
	bRecalculEnCours = False
End Sub

' Classeur fermé : oublie ses liens
Sub DonneesCarte_disposing(oEvenement As Object)
	Dim i As Long, k As Long
	On Error Resume Next
	For i = nLiens - 1 To 0 Step -1
		If EqualUnoObjects(aLiens(i)(1), oEvenement.Source) Then
			For k = i To nLiens - 2
				aLiens(k) = aLiens(k + 1)
			Next k
			nLiens = nLiens - 1
		End If
	Next i
End Sub

' Recolore la carte d'un lien avec les données actuelles. Si la carte a été
' supprimée de la feuille, le lien est retiré.
Sub RecalculerLien(oDocCalc As Object, sNom As String)
	Dim oPlage As Object, aChamps As Variant, oForme As Object, aData As Variant
	Dim aPalette As Variant, aPalettes As Variant, oUndo As Object, oColl As Object
	Dim oDoc As Object, oPage As Object, vTexte As Variant
	Dim aProps(1) As New com.sun.star.beans.PropertyValue
	On Error GoTo Erreur
	oPlage = oDocCalc.NamedRanges.getByName(sNom).getReferredCells()
	aChamps = Split(oDocCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sNom), "|")
	If UBound(aChamps) < 4 Then Exit Sub
	oForme = TrouverForme(oPlage.getSpreadsheet().getDrawPage(), aChamps(1))
	If IsNull(oForme) Then
		OublierLien(oDocCalc, sNom)
		Exit Sub
	End If
	aPalettes = ChargerPalettes()
	aPalette = aPalettes(IndicePalette(aPalettes, aChamps(2)))
	oPlage = PlageDeDonnees(oPlage)
	If IsNull(oPlage) Then Exit Sub
	aData = oPlage.getDataArray()
	' Toujours un contexte de texte : un texte désactivé par la configuration
	' est ainsi effacé
	vTexte = ContexteTexte(oDocCalc, oPlage, aData, OptionsDuLien(aChamps), oDocCalc)

	' Le recoloriage n'entre pas dans l'historique d'annulation : Ctrl+Z annule
	' la saisie de l'utilisateur, pas les changements de couleur
	oUndo = oDocCalc.getUndoManager()
	oUndo.enterHiddenUndoContext()
	oDocCalc.lockControllers()
	If aChamps(4) = "1" Then
		' Image (carte détaillée) : nouveau rendu depuis le fond de carte
		aProps(0).Name = "Hidden"
		aProps(0).Value = True
		aProps(1).Name = "AsTemplate"
		aProps(1).Value = True
		oDoc = StarDesktop.loadComponentFromURL(aChamps(3), "_blank", 0, aProps())
		vTexte(T_DOCUMENT) = oDoc
		ColorierDocument(oDoc, aData, aPalette, vTexte)
		oPage = oDoc.getDrawPages().getByIndex(0)
		oForme.Graphic = ImageDePage(oPage)
		oDoc.close(True)
	ElseIf oForme.supportsService("com.sun.star.drawing.GroupShape") Then
		ColorierConteneurs(Array(oForme), aData, aPalette, vTexte)
	Else
		oColl = CreateUnoService("com.sun.star.drawing.ShapeCollection")
		oColl.add(oForme)
		ColorierConteneurs(Array(oColl), aData, aPalette, vTexte)
	End If
	oDocCalc.unlockControllers()
	oUndo.leaveUndoContext()
	Exit Sub
Erreur:
	Journal("ERREUR au recalcul de " & sNom & " " & Err & " (ligne " & Erl & ") : " & Error$)
	On Error Resume Next
	If Not IsNull(oDoc) Then oDoc.close(True)
	If Not IsNull(oUndo) Then oUndo.leaveUndoContext()
	If oDocCalc.hasControllersLocked() Then oDocCalc.unlockControllers()
End Sub

' ---------------------------------------------------------------------------
' Configuration d'une carte insérée : clic droit sur la carte ->
' « Configuration de la carte… ». Le menu contextuel de la fenêtre du
' classeur est intercepté (une fois par fenêtre) dès qu'une carte liée y est
' suivie.
' ---------------------------------------------------------------------------
Sub IntercepterMenu(oDocCalc As Object)
	Dim oCtrl As Object, oIntercepteur As Object, i As Long
	On Error GoTo Erreur
	oCtrl = oDocCalc.getCurrentController()
	If IsNull(oCtrl) Then Exit Sub
	For i = 0 To nMenus - 1
		If EqualUnoObjects(aMenus(i)(0), oCtrl) Then Exit Sub
	Next i
	oIntercepteur = CreateUnoListener("MenuCarte_", "com.sun.star.ui.XContextMenuInterceptor")
	oCtrl.registerContextMenuInterceptor(oIntercepteur)
	ReDim Preserve aMenus(nMenus)
	aMenus(nMenus) = Array(oCtrl, oIntercepteur)
	nMenus = nMenus + 1
	Exit Sub
Erreur:
	Journal("ERREUR à l'ajout du menu contextuel : " & Error$)
End Sub

' Ajoute « Configuration de la carte… » en tête du menu si le clic droit
' porte sur une carte liée (ou une de ses zones)
Function MenuCarte_notifyContextMenuExecute(oEvenement As Object) As Variant
	Dim oSel As Object, sLien As String, oMenu As Object, oEntree As Object, oSeparateur As Object
	On Error GoTo Ignorer
	MenuCarte_notifyContextMenuExecute = com.sun.star.ui.ContextMenuInterceptorAction.IGNORED
	oSel = oEvenement.Selection.getSelection()
	' Toutes les cartes suivies sont examinées : le document n'est pas
	' toujours accessible depuis l'événement
	sLien = LienDeSelection(Nothing, oSel)
	JournalAjout("Menu contextuel : sélection " & NomImplementation(oSel) & ", carte liée '" & sLien & "'")
	If sLien = "" Then Exit Function
	' Propriétés écrites par setPropertyValue : Basic ne les voit pas comme
	' oEntree.Text (« variable d'objet non définie », LibreOffice 25.2)
	oMenu = oEvenement.ActionTriggerContainer
	oEntree = oMenu.createInstance("com.sun.star.ui.ActionTrigger")
	oEntree.setPropertyValue("Text", Tr("menu.configure"))
	oEntree.setPropertyValue("CommandURL", URL_CONFIGURER)
	' Icône : celle déclarée pour URL_CONFIGURER dans Addons.xcu (LibreOffice
	' 25.2 ignore la propriété Image de l'entrée)
	oSeparateur = oMenu.createInstance("com.sun.star.ui.ActionTriggerSeparator")
	oSeparateur.setPropertyValue("SeparatorType", com.sun.star.ui.ActionTriggerSeparatorType.LINE)
	oMenu.insertByIndex(0, oSeparateur)
	oMenu.insertByIndex(0, oEntree)
	JournalAjout("Menu : entrée « " & Tr("menu.configure") & " » ajoutée")
	MenuCarte_notifyContextMenuExecute = com.sun.star.ui.ContextMenuInterceptorAction.EXECUTE_MODIFIED
	Exit Function
Ignorer:
	JournalAjout("ERREUR dans le menu contextuel " & Err & " (ligne " & Erl & ") : " & Error$)
	MenuCarte_notifyContextMenuExecute = com.sun.star.ui.ContextMenuInterceptorAction.IGNORED
End Function

Function NomImplementation(o As Object) As String
	On Error GoTo Erreur
	NomImplementation = "(aucune)"
	If Not IsNull(o) Then NomImplementation = o.getImplementationName()
	Exit Function
Erreur:
	NomImplementation = "(inconnue)"
End Function

' Nom du lien (CalcMap_N) de la carte sélectionnée, ou d'une zone de cette
' carte ; "" si la sélection n'est pas une carte liée. oDocCalc Nothing :
' cartes de tous les documents
Function LienDeSelection(oDocCalc As Object, oSel As Object) As String
	Dim oZone As Object, oForme As Object, aChamps As Variant, i As Long, k As Long
	On Error GoTo Erreur
	LienDeSelection = ""
	If IsNull(oSel) Then Exit Function
	If HasUnoInterfaces(oSel, "com.sun.star.drawing.XShape") Then
		oZone = oSel
	ElseIf HasUnoInterfaces(oSel, "com.sun.star.drawing.XShapes") Then
		If oSel.getCount() = 0 Then Exit Function
		oZone = oSel.getByIndex(0)
	Else
		Exit Function
	End If
	For i = 0 To nLiens - 1
		If IsNull(oDocCalc) Or EqualUnoObjects(aLiens(i)(0), oDocCalc) Then
			aChamps = Split(aLiens(i)(0).getDocumentProperties().getUserDefinedProperties() _
				.getPropertyValue(aLiens(i)(3)), "|")
			oForme = TrouverForme(aLiens(i)(1).getSpreadsheet().getDrawPage(), aChamps(1))
			If Not IsNull(oForme) Then
				If EqualUnoObjects(oForme, oZone) Then
					LienDeSelection = aLiens(i)(3)
					Exit Function
				End If
				If oForme.supportsService("com.sun.star.drawing.GroupShape") Then
					For k = 0 To oForme.getCount() - 1
						If EqualUnoObjects(oForme.getByIndex(k), oZone) Then
							LienDeSelection = aLiens(i)(3)
							Exit Function
						End If
					Next k
				End If
			End If
		End If
	Next i
	Exit Function
Erreur:
	LienDeSelection = ""
End Function

' Menu « Configuration de la carte… » : rouvre le dialogue pour la carte
' sélectionnée (palette, valeurs, libellés), puis la recolore
Sub ConfigurerCarte()
	Dim oDocCalc As Object, sNom As String, aChamps As Variant, aOptions As Variant
	Dim aPalettes As Variant, aCartes As Variant, iChoix As Integer, iPalette As Integer
	Dim bInserer As Boolean, nRes As Integer, sPrefCarte As String, sPrefPalette As String
	Dim bPrefInserer As Boolean, aPrefOptions As Variant, aColonnes As Variant, oPlage As Object
	On Error GoTo Erreur
	OuvrirJournal()
	Journal("CalcMap " & VERSION & " — configuration d'une carte")
	ChargerTextes()
	oDocCalc = ThisComponent
	sNom = LienDeSelection(oDocCalc, oDocCalc.getCurrentController().getSelection())
	If sNom = "" Then
		Avertir(Tr("msg.select_map"))
		GoTo Fin
	End If
	aChamps = Split(oDocCalc.getDocumentProperties().getUserDefinedProperties().getPropertyValue(sNom), "|")
	If UBound(aChamps) < 4 Then GoTo Fin
	aOptions = OptionsDuLien(aChamps)
	oPlage = PlageDeDonnees(oDocCalc.NamedRanges.getByName(sNom).getReferredCells())
	If IsNull(oPlage) Then GoTo Fin
	aColonnes = NomsColonnes(oPlage)
	aPalettes = ChargerPalettes()
	aCartes = Array(NouvelleCarte("", NomFichier(aChamps(3)), "", aChamps(3), NomFichier(aChamps(3)), _
		0, "", True))
	bInserer = True
	nRes = DialogueCartes(aCartes, aPalettes, "", "", aChamps(2), iChoix, iPalette, bInserer, aOptions, _
		aColonnes, FenetreParente(oDocCalc), True)
	Journal("Configuration de " & sNom & " : résultat " & nRes & ", palette n° " & iPalette & ", " _
		& TexteOptions(aOptions))
	If nRes <> 1 Or iPalette < 0 Then GoTo Fin
	oDocCalc.getDocumentProperties().getUserDefinedProperties().setPropertyValue(sNom, _
		ChampsLien(aChamps(1), aPalettes(iPalette)(P_ID), aChamps(3), aChamps(4) = "1", aOptions))
	' Les choix deviennent ceux proposés par défaut pour la prochaine carte
	' (sauf les colonnes, non mémorisées)
	LirePreferences(sPrefCarte, sPrefPalette, bPrefInserer, aPrefOptions)
	EcrirePreferences(sPrefCarte, aPalettes(iPalette)(P_ID), bPrefInserer, aOptions)
	RecalculerLien(oDocCalc, sNom)
	GoTo Fin
Erreur:
	Journal("ERREUR à la configuration de la carte " & Err & " (ligne " & Erl & ") : " & Error$)
	On Error Resume Next
	Avertir(Tr("msg.partial"))
Fin:
	On Error Resume Next
	Journal("Fin.")
	FermerJournal()
End Sub

' Retire un lien dont la carte a disparu (plage nommée, propriété, écouteur)
Sub OublierLien(oDocCalc As Object, sNom As String)
	Dim i As Long, k As Long
	On Error Resume Next
	oDocCalc.getDocumentProperties().getUserDefinedProperties().removeProperty(sNom)
	oDocCalc.NamedRanges.removeByName(sNom)
	For i = nLiens - 1 To 0 Step -1
		If EqualUnoObjects(aLiens(i)(0), oDocCalc) And aLiens(i)(3) = sNom Then
			aLiens(i)(1).removeModifyListener(aLiens(i)(2))
			For k = i To nLiens - 2
				aLiens(k) = aLiens(k + 1)
			Next k
			nLiens = nLiens - 1
		End If
	Next i
End Sub

Function NomCellule(oCellule As Object) As String
	On Error GoTo Erreur
	NomCellule = oCellule.AbsoluteName
	Exit Function
Erreur:
	NomCellule = "?"
End Function

' Nom de fichier sans extension, depuis une URL
Function NomFichier(sUrl As String) As String
	Dim aParties As Variant, s As String
	aParties = Split(ConvertFromURL(sUrl), GetPathSeparator())
	s = aParties(UBound(aParties))
	If LCase(Right(s, 4)) = ".odg" Then s = Left(s, Len(s) - 4)
	NomFichier = s
End Function

' ---------------------------------------------------------------------------
' Palettes de couleurs : palettes.json de l'extension, puis celui de
' l'utilisateur (dossier du cache), qui ajoute des palettes ou remplace
' celles de même id.
' Palette -> Array(id, nom, couleurs(), couleur sans donnée, dégradé continu).
' ---------------------------------------------------------------------------
Function ChargerPalettes() As Variant
	Dim aRes() As Variant, n As Long, aLues As Variant, sDossier As String
	Dim i As Long, k As Long, bRemplacee As Boolean
	On Error GoTo Erreur
	sDossier = DossierExtension()
	If sDossier <> "" Then aRes = LirePalettes(sDossier & NOM_PALETTES)
	n = UBound(aRes) + 1
	sDossier = DossierCartes()
	If sDossier <> "" Then
		aLues = LirePalettes(sDossier & NOM_PALETTES)
		For i = 0 To UBound(aLues)
			bRemplacee = False
			For k = 0 To n - 1
				If aRes(k)(P_ID) = aLues(i)(P_ID) Then
					aRes(k) = aLues(i)
					bRemplacee = True
				End If
			Next k
			If Not bRemplacee Then
				ReDim Preserve aRes(n)
				aRes(n) = aLues(i)
				n = n + 1
			End If
		Next i
	End If
Fin:
	If n = 0 Then
		ReDim aRes(0)
		aRes(0) = NouvellePalette(PALETTE_PAR_DEFAUT, "Bleus", _
			Array(RGB(239, 243, 255), RGB(189, 215, 231), RGB(107, 174, 214), RGB(49, 130, 189), RGB(8, 81, 156)), _
			RGB(230, 230, 230), False)
	End If
	Journal("Palettes disponibles : " & UBound(aRes) + 1)
	ChargerPalettes = aRes
	Exit Function
Erreur:
	Journal("ERREUR au chargement des palettes : " & Error$)
	Resume Fin
End Function

' Lit un fichier de palettes ; ignore les palettes incomplètes (journal)
Function LirePalettes(sUrl As String) As Variant
	Dim oRacine As Variant, aListe As Variant, oPal As Variant, aTeintes As Variant
	Dim aRes() As Variant, aCouleurs() As Long, n As Long, i As Long, k As Long
	Dim sId As String, nVide As Long, bValide As Boolean, vContinu As Variant
	On Error GoTo Erreur
	LirePalettes = Array()
	If Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(sUrl) Then Exit Function
	oRacine = JsonLire(LireTexte(sUrl))
	aListe = JsonElements(JsonChamp(oRacine, "palettes"))
	For i = 0 To UBound(aListe)
		oPal = aListe(i)
		sId = JsonTexte(oPal, "id")
		aTeintes = JsonElements(JsonChamp(oPal, "colors"))
		bValide = (sId <> "" And UBound(aTeintes) >= 1)
		If bValide Then
			ReDim aCouleurs(UBound(aTeintes)) As Long
			For k = 0 To UBound(aTeintes)
				aCouleurs(k) = CouleurHexa(aTeintes(k))
				If aCouleurs(k) < 0 Then bValide = False
			Next k
		End If
		If bValide Then
			nVide = CouleurHexa(JsonTexte(oPal, "nodata"))
			If nVide < 0 Then nVide = RGB(230, 230, 230)
			vContinu = JsonChamp(oPal, "continuous")
			ReDim Preserve aRes(n)
			aRes(n) = NouvellePalette(sId, NomPalette(oPal, sId), aCouleurs, nVide, _
				(VarType(vContinu) = 11) And (vContinu = True))
			n = n + 1
		Else
			Journal("Palette " & (i + 1) & " de " & ConvertFromURL(sUrl) _
				& " ignorée : id manquant, moins de 2 couleurs ou couleur invalide")
		End If
	Next i
	If n > 0 Then LirePalettes = aRes
	Exit Function
Erreur:
	Journal("ERREUR à la lecture des palettes " & ConvertFromURL(sUrl) & " : " & Error$)
	If n > 0 Then LirePalettes = aRes
End Function

Function NouvellePalette(sId, sNom, aCouleurs, nVide, bContinu) As Variant
	NouvellePalette = Array(sId, sNom, aCouleurs, nVide, bContinu)
End Function

' Nom de palette : texte simple, ou objet {"fr": ..., "en": ...}
Function NomPalette(oPal As Variant, sId As String) As String
	Dim v As Variant, s As String
	v = JsonChamp(oPal, "name")
	If VarType(v) = 8 Then
		s = v
	Else
		s = JsonTexte(v, Langue())
		If s = "" Then s = JsonTexte(v, "en")
	End If
	If s = "" Then s = sId
	NomPalette = s
End Function

' Indice de la palette sId, sinon de la palette par défaut, sinon 0
Function IndicePalette(aPalettes As Variant, sId As String) As Integer
	Dim i As Integer
	IndicePalette = 0
	For i = 0 To UBound(aPalettes)
		If aPalettes(i)(P_ID) = PALETTE_PAR_DEFAUT Then IndicePalette = i
	Next i
	For i = 0 To UBound(aPalettes)
		If sId <> "" And aPalettes(i)(P_ID) = sId Then IndicePalette = i
	Next i
End Function

' "#rrggbb" -> couleur ; -1 si invalide
Function CouleurHexa(v As Variant) As Long
	Dim s As String, k As Integer, nChiffre As Integer, nValeur As Long
	CouleurHexa = -1
	If VarType(v) <> 8 Then Exit Function
	s = LCase(Trim(v))
	If Left(s, 1) = "#" Then s = Mid(s, 2)
	If Len(s) <> 6 Then Exit Function
	For k = 1 To 6
		nChiffre = InStr(1, "0123456789abcdef", Mid(s, k, 1), 0) - 1
		If nChiffre < 0 Then Exit Function
		nValeur = nValeur * 16 + nChiffre
	Next k
	CouleurHexa = nValeur
End Function

' ---------------------------------------------------------------------------
' Traductions : i18n/messages_<langue>.properties de l'extension. L'anglais
' est chargé d'abord puis complété par la langue de l'interface.
' ---------------------------------------------------------------------------
Sub ChargerTextes()
	Dim sDossier As String
	On Error GoTo Erreur
	oTextes = com.sun.star.container.EnumerableMap.create("string", "string")
	sDossier = DossierExtension()
	If sDossier = "" Then Exit Sub
	LireProprietes(sDossier & "i18n/messages_en.properties", oTextes)
	If Langue() <> "en" Then LireProprietes(sDossier & "i18n/messages_" & Langue() & ".properties", oTextes)
	Journal("Langue de l'interface : " & Langue())
	Exit Sub
Erreur:
	Journal("ERREUR au chargement des traductions : " & Error$)
End Sub

' Texte traduit ; %1 et %2 sont remplacés par v1 et v2. Clé inconnue : la clé.
Function Tr(sCle As String, Optional v1 As Variant, Optional v2 As Variant) As String
	Dim s As String
	On Error GoTo Erreur
	If IsNull(oTextes) Then ChargerTextes()
	s = sCle
	If oTextes.containsKey(sCle) Then s = oTextes.get(sCle)
	If Not IsMissing(v1) Then s = Replace(s, "%1", CStr(v1))
	If Not IsMissing(v2) Then s = Replace(s, "%2", CStr(v2))
	Tr = s
	Exit Function
Erreur:
	Tr = sCle
End Function

' Code de langue de l'interface ("fr", "en"…)
Function Langue() As String
	Dim s As String
	If sLangue = "" Then
		s = LireConfiguration("/org.openoffice.Office.Linguistic/General", "UILocale")
		If s = "" Then s = LireConfiguration("/org.openoffice.Setup/L10N", "ooLocale")
		s = LCase(Left(s, 2))
		If s = "" Then s = "en"
		sLangue = s
	End If
	Langue = sLangue
End Function

Function LireConfiguration(sNoeud As String, sPropriete As String) As String
	Dim oFournisseur As Object, oAcces As Object
	Dim aArgs(0) As New com.sun.star.beans.PropertyValue
	On Error GoTo Erreur
	oFournisseur = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
	aArgs(0).Name = "nodepath"
	aArgs(0).Value = sNoeud
	oAcces = oFournisseur.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", aArgs())
	LireConfiguration = oAcces.getByName(sPropriete)
	Exit Function
Erreur:
	LireConfiguration = ""
End Function

' Dossier d'installation de l'extension ("" si introuvable)
Function DossierExtension() As String
	Dim s As String
	On Error GoTo Erreur
	s = GetDefaultContext().getByName("/singletons/com.sun.star.deployment.PackageInformationProvider") _
		.getPackageLocation(ID_EXTENSION)
	If s <> "" Then s = s & "/"
	DossierExtension = s
	Exit Function
Erreur:
	DossierExtension = ""
End Function

' Lit un fichier cle=valeur (UTF-8, # commentaire, \n retour à la ligne)
' et ajoute ses entrées à oTable (EnumerableMap string -> string)
Sub LireProprietes(sUrl As String, oTable As Object)
	Dim aLignes As Variant, sLigne As String, i As Long, p As Long
	On Error GoTo Erreur
	If Not CreateUnoService("com.sun.star.ucb.SimpleFileAccess").exists(sUrl) Then Exit Sub
	aLignes = Split(LireTexte(sUrl), Chr(10))
	For i = 0 To UBound(aLignes)
		sLigne = LTrim(Replace(aLignes(i), Chr(13), ""))
		p = InStr(sLigne, "=")
		If p > 1 And Left(sLigne, 1) <> "#" Then
			oTable.put(Trim(Left(sLigne, p - 1)), Replace(LTrim(Mid(sLigne, p + 1)), "\n", Chr(10)))
		End If
	Next i
	Exit Sub
Erreur:
	Journal("ERREUR à la lecture de " & ConvertFromURL(sUrl) & " : " & Error$)
End Sub

Sub EcrireTexte(sUrl As String, sTexte As String)
	Dim oSFA As Object, oFlux As Object
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sUrl) Then oSFA.kill(sUrl)
	oFlux = CreateUnoService("com.sun.star.io.TextOutputStream")
	oFlux.setOutputStream(oSFA.openFileWrite(sUrl))
	oFlux.setEncoding("UTF-8")
	oFlux.writeString(sTexte)
	oFlux.closeOutput()
End Sub

' ---------------------------------------------------------------------------
' Lecteur JSON minimal. Objet -> Array("{", clés(), valeurs()) ;
' tableau -> Array("[", éléments()). Empty si le texte est invalide.
' ---------------------------------------------------------------------------
Function JsonLire(s As String) As Variant
	sJson = s
	nJson = 1
	bJsonOk = True
	JsonLire = JsonValeur()
	JsonBlancs()
	If nJson <= Len(sJson) Then bJsonOk = False
	If Not bJsonOk Then
		Journal("JSON invalide près du caractère " & nJson)
		JsonLire = Empty
	End If
	sJson = ""
End Function

Function JsonChamp(o As Variant, sCle As String) As Variant
	Dim aCles As Variant, i As Long
	JsonChamp = Empty
	If Not IsArray(o) Then Exit Function
	If o(0) <> "{" Then Exit Function
	aCles = o(1)
	For i = 0 To UBound(aCles)
		If aCles(i) = sCle Then
			JsonChamp = o(2)(i)
			Exit Function
		End If
	Next i
End Function

Function JsonElements(o As Variant) As Variant
	JsonElements = Array()
	If Not IsArray(o) Then Exit Function
	If o(0) = "[" Then JsonElements = o(1)
End Function

Function JsonTexte(o As Variant, sCle As String) As String
	Dim v As Variant
	v = JsonChamp(o, sCle)
	If VarType(v) = 8 Then JsonTexte = v Else JsonTexte = ""
End Function

' Champ traduit s'il existe (ex. "title_en"), sinon le champ de base
Function JsonTexteTraduit(o As Variant, sCle As String) As String
	Dim s As String
	s = JsonTexte(o, sCle & "_" & Langue())
	If s = "" Then s = JsonTexte(o, sCle)
	JsonTexteTraduit = s
End Function

Function JsonNombre(o As Variant, sCle As String) As Double
	Dim v As Variant
	v = JsonChamp(o, sCle)
	If VarType(v) = 5 Then JsonNombre = v Else JsonNombre = 0
End Function

Sub JsonBlancs()
	Dim c As String
	Do While nJson <= Len(sJson)
		c = Mid(sJson, nJson, 1)
		If c <> " " And c <> Chr(9) And c <> Chr(10) And c <> Chr(13) Then Exit Do
		nJson = nJson + 1
	Loop
End Sub

Function JsonValeur() As Variant
	Dim c As String
	JsonValeur = Empty
	JsonBlancs()
	If nJson > Len(sJson) Then
		bJsonOk = False
		Exit Function
	End If
	c = Mid(sJson, nJson, 1)
	If c = "{" Then
		JsonValeur = JsonObjet()
	ElseIf c = "[" Then
		JsonValeur = JsonTableau()
	ElseIf c = """" Then
		JsonValeur = JsonChaine()
	ElseIf c = "t" Then
		JsonValeur = JsonMot("true", True)
	ElseIf c = "f" Then
		JsonValeur = JsonMot("false", False)
	ElseIf c = "n" Then
		JsonValeur = JsonMot("null", Null)
	Else
		JsonValeur = JsonNum()
	End If
End Function

Function JsonObjet() As Variant
	Dim aCles() As Variant, aValeurs() As Variant, n As Long, c As String
	nJson = nJson + 1
	JsonBlancs()
	If Mid(sJson, nJson, 1) = "}" Then
		nJson = nJson + 1
		JsonObjet = Array("{", aCles, aValeurs)
		Exit Function
	End If
	Do
		JsonBlancs()
		If Mid(sJson, nJson, 1) <> """" Then
			bJsonOk = False
			Exit Function
		End If
		ReDim Preserve aCles(n)
		ReDim Preserve aValeurs(n)
		aCles(n) = JsonChaine()
		JsonBlancs()
		If Mid(sJson, nJson, 1) <> ":" Then
			bJsonOk = False
			Exit Function
		End If
		nJson = nJson + 1
		aValeurs(n) = JsonValeur()
		If Not bJsonOk Then Exit Function
		n = n + 1
		JsonBlancs()
		c = Mid(sJson, nJson, 1)
		nJson = nJson + 1
		If c = "}" Then Exit Do
		If c <> "," Then
			bJsonOk = False
			Exit Function
		End If
	Loop
	JsonObjet = Array("{", aCles, aValeurs)
End Function

Function JsonTableau() As Variant
	Dim aElements() As Variant, n As Long, c As String
	nJson = nJson + 1
	JsonBlancs()
	If Mid(sJson, nJson, 1) = "]" Then
		nJson = nJson + 1
		JsonTableau = Array("[", aElements)
		Exit Function
	End If
	Do
		ReDim Preserve aElements(n)
		aElements(n) = JsonValeur()
		If Not bJsonOk Then Exit Function
		n = n + 1
		JsonBlancs()
		c = Mid(sJson, nJson, 1)
		nJson = nJson + 1
		If c = "]" Then Exit Do
		If c <> "," Then
			bJsonOk = False
			Exit Function
		End If
	Loop
	JsonTableau = Array("[", aElements)
End Function

Function JsonChaine() As String
	Dim s As String, c As String, k As Integer, nCode As Long
	nJson = nJson + 1
	Do While nJson <= Len(sJson)
		c = Mid(sJson, nJson, 1)
		If c = """" Then
			nJson = nJson + 1
			JsonChaine = s
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
				' caractères de contrôle ignorés
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
	JsonChaine = s
End Function

Function JsonMot(sMot As String, v As Variant) As Variant
	If Mid(sJson, nJson, Len(sMot)) = sMot Then
		nJson = nJson + Len(sMot)
		JsonMot = v
	Else
		bJsonOk = False
		JsonMot = Empty
	End If
End Function

Function JsonNum() As Double
	Dim nDebut As Long
	nDebut = nJson
	Do While nJson <= Len(sJson)
		If InStr(1, "+-0123456789.eE", Mid(sJson, nJson, 1), 0) = 0 Then Exit Do
		nJson = nJson + 1
	Loop
	If nJson = nDebut Then
		bJsonOk = False
		JsonNum = 0
	Else
		JsonNum = Val(Mid(sJson, nDebut, nJson - nDebut))
	End If
End Function

' ---------------------------------------------------------------------------
' Messages et journal
' ---------------------------------------------------------------------------
Sub Avertir(sMessage As String)
	On Error Resume Next
	Journal("Message : " & sMessage)
	If sCheminJournal <> "" Then
		MsgBox sMessage & Chr(10) & Chr(10) & Tr("msg.details", sCheminJournal), 48, Tr("app.title")
	Else
		MsgBox sMessage, 48, Tr("app.title")
	End If
End Sub

Sub Informer(sMessage As String)
	On Error Resume Next
	Journal("Information : " & sMessage)
	MsgBox sMessage, 64, Tr("app.title")
End Sub

Sub OuvrirJournal()
	Dim oSFA As Object, oFlux As Object, sUrl As String
	On Error GoTo Erreur
	oJournal = Nothing
	sCheminJournal = ""
	sUrl = CreateUnoService("com.sun.star.util.PathSubstitution") _
		.substituteVariables("$(user)/" & NOM_JOURNAL, True)
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	If oSFA.exists(sUrl) Then oSFA.kill(sUrl)
	oFlux = oSFA.openFileWrite(sUrl)
	oJournal = CreateUnoService("com.sun.star.io.TextOutputStream")
	oJournal.setOutputStream(oFlux)
	oJournal.setEncoding("UTF-8")
	sCheminJournal = ConvertFromURL(sUrl)
	Exit Sub
Erreur:
	oJournal = Nothing        ' pas de journal : la macro fonctionne quand même
End Sub

' Ajoute une ligne au journal en dehors d'une exécution de la macro
' (écouteurs : menu contextuel…)
Sub JournalAjout(sTexte As String)
	Dim oSFA As Object, oFlux As Object, oTexte As Object, sUrl As String
	On Error GoTo Erreur
	If Not IsNull(oJournal) Then
		Journal(sTexte)
		Exit Sub
	End If
	sUrl = CreateUnoService("com.sun.star.util.PathSubstitution") _
		.substituteVariables("$(user)/" & NOM_JOURNAL, True)
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	oFlux = oSFA.openFileReadWrite(sUrl)
	oFlux.seek(oFlux.getLength())
	oTexte = CreateUnoService("com.sun.star.io.TextOutputStream")
	oTexte.setOutputStream(oFlux.getOutputStream())
	oTexte.setEncoding("UTF-8")
	oTexte.writeString(Format(Now, "yyyy-mm-dd hh:mm:ss") & "  " & sTexte & Chr(10))
	oTexte.closeOutput()
Erreur:
End Sub

Sub Journal(sTexte As String)
	On Error Resume Next
	If IsNull(oJournal) Then Exit Sub
	oJournal.writeString(Format(Now, "yyyy-mm-dd hh:mm:ss") & "  " & sTexte & Chr(10))
End Sub

Sub FermerJournal()
	On Error Resume Next
	If Not IsNull(oJournal) Then oJournal.closeOutput()
	oJournal = Nothing
End Sub
