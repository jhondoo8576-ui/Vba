Option Explicit

'========================================================
' PORTAIL NAVAL GROUP - VERSION 1.0
' VBA NATIF - MODULE STANDARD
' Aucune instruction WITH
'========================================================

Private Const NOM_FEUILLE As String = "TEST"
Private Const FEUILLE_IMPORT As String = "DONNEES_IMPORT"

Private Const BLEU_MARINE As Long = 394758
Private Const BLEU_NAVAL As Long = 12544704
Private Const BLEU_CLAIR As Long = 15722479
Private Const GRIS_FOND As Long = 15987699
Private Const BLANC As Long = 16777215
Private Const GRIS_TEXTE As Long = 7566195

Private Const CELLULE_FICHIER As String = "B23"

'========================================================
' 1. GENERATION GENERALE
'========================================================

Public Sub GenererInterface()

    Dim ws As Worksheet
    Dim ancienneMaj As Boolean

    ancienneMaj = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Application.ScreenUpdating = False

    If FeuilleExiste(NOM_FEUILLE) = False Then
        MsgBox "La feuille TEST est introuvable.", vbExclamation
        GoTo Sortie
    End If

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE)

    ConfigurerFeuille ws

    CreerTexte ws, "TitreNG", "NAVAL GROUP", _
        35, 22, 250, 30, 20, BLANC, True

    CreerTexte ws, "LogoNG", "NG", _
        925, 18, 55, 48, 23, BLANC, True

    CreerTexte ws, "Application", "PORTAIL DE GESTION", _
        35, 52, 400, 25, 10, BLEU_CLAIR, True

    CreerTexte ws, "NomApplication", "ESPACE DE GESTION DOCUMENTAIRE", _
        300, 25, 570, 35, 17, BLANC, True

    ' MENU PRINCIPAL

    CreerBouton ws, "BtnAccueil", "TABLEAU DE BORD", _
        35, 110, 145, 45, "AllerTableauBord", BLEU_MARINE

    CreerBouton ws, "BtnImport", "IMPORTER UN FICHIER", _
        190, 110, 145, 45, "OuvrirImportation", BLEU_NAVAL

    CreerBouton ws, "BtnNouveau", "NOUVEAU DOSSIER", _
        345, 110, 130, 45, "NouveauDossier", BLEU_MARINE

    CreerBouton ws, "BtnConsultation", "CONSULTATION", _
        485, 110, 125, 45, "OuvrirConsultation", BLEU_MARINE

    CreerBouton ws, "BtnRapports", "RAPPORTS", _
        620, 110, 110, 45, "OuvrirRapports", BLEU_MARINE

    CreerBouton ws, "BtnParametres", "PARAMÈTRES", _
        740, 110, 115, 45, "OuvrirParametres", BLEU_MARINE

    CreerBouton ws, "BtnAssistance", "ASSISTANCE", _
        865, 110, 115, 45, "OuvrirAssistance", BLEU_MARINE

    ' TABLEAU DE BORD

    CreerTexte ws, "TitreDashboard", "TABLEAU DE BORD", _
        40, 185, 500, 40, 25, BLEU_MARINE, True

    CreerTexte ws, "SousTitreDashboard", _
        "Bienvenue dans votre espace de gestion documentaire.", _
        40, 225, 700, 25, 11, GRIS_TEXTE, False

    ' CARTES

    CreerCarte ws, "Carte1", "ESPACE DE TRAVAIL", _
        "Interface de gestion centralisée", _
        40, 280, 285, 115

    CreerCarte ws, "Carte2", "IMPORTATION", _
        "Sélection et intégration de fichiers", _
        355, 280, 285, 115

    CreerCarte ws, "Carte3", "ENVIRONNEMENT", _
        "Version applicative : 1.0", _
        670, 280, 285, 115

    CreerTexte ws, "ValeurImport", "AUCUNE IMPORTATION", _
        375, 355, 245, 22, 10, BLEU_NAVAL, True

    ' ZONE IMPORTATION

    CreerZoneImportation ws

    ' FOOTER

    CreerTexte ws, "Footer", _
        "NAVAL GROUP     |     PORTAIL DE GESTION     |     VERSION 1.0", _
        35, 665, 700, 22, 9, GRIS_TEXTE, True

    CreerTexte ws, "FooterStatut", _
        "ENVIRONNEMENT DE TRAVAIL", _
        750, 665, 230, 22, 9, BLEU_NAVAL, True

    CreerListeFichiers

    ActualiserTableauBord

Sortie:

    Application.ScreenUpdating = ancienneMaj

    Exit Sub

GestionErreur:

    MsgBox "Erreur lors de la création du portail : " & _
        Err.Description, vbCritical

    Resume Sortie

End Sub

'========================================================
' 2. CONFIGURATION DE LA FEUILLE
'========================================================

Public Sub ConfigurerFeuille(ByVal ws As Worksheet)

    Dim i As Long

    ws.Activate

    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.Zoom = 85

    ws.Cells.Interior.Color = GRIS_FOND
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ws.Rows.RowHeight = 22

    ws.Columns("A:Z").ColumnWidth = 12

    ' Nettoyage des anciennes formes

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    ws.Range("AA:AB").EntireColumn.Hidden = True

End Sub

'========================================================
' 3. CREATION DES TEXTES
'========================================================

Public Sub CreerTexte( _
    ByVal ws As Worksheet, _
    ByVal nom As String, _
    ByVal contenu As String, _
    ByVal gauche As Double, _
    ByVal haut As Double, _
    ByVal largeur As Double, _
    ByVal hauteur As Double, _
    ByVal taille As Double, _
    ByVal couleur As Long, _
    ByVal gras As Boolean)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRectangle, gauche, haut, largeur, hauteur)

    shp.Name = nom

    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse

    shp.TextFrame2.TextRange.Text = contenu
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = taille
    shp.TextFrame2.TextRange.Font.Bold = gras
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = couleur

    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

    shp.TextFrame2.MarginLeft = 2
    shp.TextFrame2.MarginRight = 2

End Sub

'========================================================
' 4. CREATION DES BOUTONS
'========================================================

Public Sub CreerBouton( _
    ByVal ws As Worksheet, _
    ByVal nom As String, _
    ByVal texte As String, _
    ByVal gauche As Double, _
    ByVal haut As Double, _
    ByVal largeur As Double, _
    ByVal hauteur As Double, _
    ByVal macro As String, _
    ByVal couleur As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        gauche, haut, largeur, hauteur)

    shp.Name = nom

    shp.Fill.ForeColor.RGB = couleur
    shp.Line.Visible = msoFalse

    shp.TextFrame2.TextRange.Text = texte
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 9
    shp.TextFrame2.TextRange.Font.Bold = True
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BLANC

    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    shp.TextFrame2.MarginLeft = 5
    shp.TextFrame2.MarginRight = 5

    shp.OnAction = "'" & ThisWorkbook.Name & "'!" & macro

End Sub

'========================================================
' 5. CREATION DES CARTES
'========================================================

Public Sub CreerCarte( _
    ByVal ws As Worksheet, _
    ByVal nom As String, _
    ByVal titre As String, _
    ByVal description As String, _
    ByVal gauche As Double, _
    ByVal haut As Double, _
    ByVal largeur As Double, _
    ByVal hauteur As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        gauche, haut, largeur, hauteur)

    shp.Name = nom

    shp.Fill.ForeColor.RGB = BLANC
    shp.Line.ForeColor.RGB = BLEU_CLAIR
    shp.Line.Weight = 1

    CreerTexte ws, nom & "Titre", titre, _
        gauche + 18, haut + 18, largeur - 35, 25, _
        12, BLEU_MARINE, True

    CreerTexte ws, nom & "Description", description, _
        gauche + 18, haut + 55, largeur - 35, 35, _
        10, GRIS_TEXTE, False

End Sub

'========================================================
' 6. ZONE D'IMPORTATION
'========================================================

Public Sub CreerZoneImportation(ByVal ws As Worksheet)

    Dim fond As Shape
    Dim position As Double

    position = 435

    Set fond = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        40, position, 915, 195)

    fond.Name = "ZoneImportation"
    fond.Fill.ForeColor.RGB = BLANC
    fond.Line.ForeColor.RGB = BLEU_CLAIR
    fond.Line.Weight = 1

    CreerTexte ws, "TitreImportation", _
        "SÉLECTION DU FICHIER À IMPORTER", _
        60, 455, 600, 30, 15, BLEU_MARINE, True

    CreerTexte ws, "DescriptionImportation", _
        "Choisissez un fichier Excel disponible ou parcourez vos dossiers.", _
        60, 490, 750, 25, 10, GRIS_TEXTE, False

    ' Zone de sélection Excel

    ws.Range("B23:F23").Merge

    ws.Range("B23").Interior.Color = BLANC
    ws.Range("B23").Font.Name = "Segoe UI"
    ws.Range("B23").Font.Size = 10
    ws.Range("B23").Font.Color = BLEU_MARINE

    ws.Range("B23:F23").Borders.LineStyle = xlContinuous
    ws.Range("B23:F23").Borders.Color = BLEU_NAVAL

    CreerBouton ws, "BtnParcourir", "PARCOURIR", _
        550, 520, 150, 38, "ParcourirFichier", BLEU_NAVAL

    CreerBouton ws, "BtnImporter", "IMPORTER", _
        715, 520, 150, 38, "ImporterFichier", BLEU_MARINE

    CreerTexte ws, "StatutImport", _
        "En attente de sélection du fichier.", _
        60, 575, 750, 25, 10, BLEU_NAVAL, False

End Sub

'========================================================
' 7. CREATION DE LA LISTE DEROULANTE
'========================================================

Public Sub CreerListeFichiers()

    Dim ws As Worksheet
    Dim dossier As String
    Dim fichier As String
    Dim derniereLigne As Long

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE)

    ws.Range("AA:AB").ClearContents

    ws.Range("AA1").Value = "Aucun fichier disponible"
    ws.Range("AB1").Value = ""

    If Len(ThisWorkbook.Path) > 0 Then

        dossier = ThisWorkbook.Path & Application.PathSeparator
        fichier = Dir(dossier & "*.xls*")

        Do While Len(fichier) > 0

            If EstFichierExcel(fichier) Then

                If fichier <> ThisWorkbook.Name Then

                    AjouterFichierListe dossier & fichier

                End If

            End If

            fichier = Dir()

        Loop

    End If

    derniereLigne = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    If derniereLigne < 1 Then
        derniereLigne = 1
    End If

    On Error Resume Next
    ws.Range(CELLULE_FICHIER).Validation.Delete
    On Error GoTo 0

    ws.Range(CELLULE_FICHIER).Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="=$AA$1:$AA$" & derniereLigne

    ws.Range(CELLULE_FICHIER).Validation.IgnoreBlank = True
    ws.Range(CELLULE_FICHIER).Validation.InCellDropdown = True

    ws.Range(CELLULE_FICHIER).Value = ""

End Sub

'========================================================
' 8. AJOUT D'UN FICHIER DANS LA LISTE
'========================================================

Private Sub AjouterFichierListe(ByVal chemin As String)

    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim nomFichier As String
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE)

    nomFichier = Mid$(chemin, InStrRev(chemin, _
        Application.PathSeparator) + 1)

    derniereLigne = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    If ws.Range("AA1").Value = "Aucun fichier disponible" Then
        derniereLigne = 0
        ws.Range("AA1:AB1").ClearContents
    End If

    For i = 1 To derniereLigne

        If StrComp(CStr(ws.Cells(i, "AB").Value), _
            chemin, vbTextCompare) = 0 Then

            Exit Sub

        End If

    Next i

    derniereLigne = derniereLigne + 1

    ws.Cells(derniereLigne, "AA").Value = nomFichier
    ws.Cells(derniereLigne, "AB").Value = chemin

    CreerValidationFichiers ws, derniereLigne

End Sub

Private Sub CreerValidationFichiers( _
    ByVal ws As Worksheet, _
    ByVal derniereLigne As Long)

    On Error Resume Next
    ws.Range(CELLULE_FICHIER).Validation.Delete
    On Error GoTo 0

    ws.Range(CELLULE_FICHIER).Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=$AA$1:$AA$" & derniereLigne

    ws.Range(CELLULE_FICHIER).Validation.InCellDropdown = True

End Sub

'========================================================
' 9. PARCOURIR UN FICHIER
'========================================================

Public Sub ParcourirFichier()

    Dim boite As FileDialog
    Dim chemin As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE)

    Set boite = Application.FileDialog(msoFileDialogFilePicker)

    boite.Title = "Sélectionner un fichier Excel"
    boite.AllowMultiSelect = False

    boite.Filters.Clear
    boite.Filters.Add "Fichiers Excel", "*.xlsx;*.xlsm;*.xls"

    If boite.Show <> -1 Then

        ws.Shapes("StatutImport").TextFrame2.TextRange.Text = _
            "Sélection annulée."

        Exit Sub

    End If

    chemin = boite.SelectedItems(1)

    If EstFichierExcel(chemin) = False Then

        MsgBox "Format de fichier non accepté.", vbExclamation
        Exit Sub

    End If

    AjouterFichierListe chemin

    ws.Range(CELLULE_FICHIER).Value = _
        Mid$(chemin, InStrRev(chemin, _
        Application.PathSeparator) + 1)

    ws.Shapes("StatutImport").TextFrame2.TextRange.Text = _
        "Fichier sélectionné : " & ws.Range(CELLULE_FICHIER).Value

End Sub

'========================================================
' 10. IMPORTATION DU FICHIER
'========================================================

Public Sub ImporterFichier()

    Dim ws As Worksheet
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet

    Dim classeurSource As Workbook
    Dim plageSource As Range

    Dim chemin As String
    Dim ancienneMaj As Boolean

    Dim erreur As String

    ancienneMaj = Application.ScreenUpdating

    On Error GoTo GestionErreur

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE)

    If Len(Trim$(CStr(ws.Range(CELLULE_FICHIER).Value))) = 0 Then

        MsgBox "Veuillez sélectionner un fichier.", vbExclamation
        Exit Sub

    End If

    chemin = ObtenirCheminFichier( _
        CStr(ws.Range(CELLULE_FICHIER).Value))

    If Len(chemin) = 0 Then

        MsgBox "Impossible de retrouver le chemin du fichier.", _
            vbExclamation

        Exit Sub

    End If

    If Len(Dir(chemin)) = 0 Then

        MsgBox "Le fichier sélectionné n'existe plus.", vbExclamation
        Exit Sub

    End If

    Application.ScreenUpdating = False

    Set classeurSource = Workbooks.Open( _
        Filename:=chemin, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        AddToMru:=False)

    Set wsSource = classeurSource.Worksheets(1)
    Set plageSource = wsSource.UsedRange

    If FeuilleExiste(FEUILLE_IMPORT) Then

        Set wsDestination = ThisWorkbook.Worksheets(FEUILLE_IMPORT)

    Else

        Set wsDestination = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.Count))

        wsDestination.Name = FEUILLE_IMPORT

    End If

    If Application.WorksheetFunction.CountA( _
        wsDestination.Cells) > 0 Then

        If MsgBox( _
            "La feuille DONNEES_IMPORT contient déjà des données." & _
            vbCrLf & vbCrLf & _
            "Voulez-vous les remplacer ?", _
            vbYesNo + vbQuestion, _
            "Confirmation d'importation") <> vbYes Then

            classeurSource.Close SaveChanges:=False
            Set classeurSource = Nothing

            GoTo Sortie

        End If

    End If

    wsDestination.Cells.Clear

    ' IMPORTATION DES DONNEES ET DES FORMATS

    plageSource.Copy Destination:=wsDestination.Range("A1")

    Application.CutCopyMode = False

    classeurSource.Close SaveChanges:=False
    Set classeurSource = Nothing

    ws.Shapes("StatutImport").TextFrame2.TextRange.Text = _
        "Importation réalisée avec succès."

    MsgBox "Le fichier a été importé avec succès dans " & _
        FEUILLE_IMPORT & ".", vbInformation

    ActualiserTableauBord

Sortie:

    Application.CutCopyMode = False
    Application.ScreenUpdating = ancienneMaj

    Exit Sub

GestionErreur:

    erreur = Err.Description

    If Not classeurSource Is Nothing Then

        On Error Resume Next
        classeurSource.Close SaveChanges:=False
        Set classeurSource = Nothing
        On Error GoTo 0

    End If

    Application.CutCopyMode = False
    Application.ScreenUpdating = ancienneMaj

    MsgBox "Erreur pendant l'importation : " & erreur, vbCritical

End Sub

'========================================================
' 11. RETROUVER LE CHEMIN COMPLET
'========================================================

Private Function ObtenirCheminFichier( _
    ByVal nomFichier As String) As String

    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE)

    derniereLigne = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    For i = 1 To derniereLigne

        If StrComp( _
            CStr(ws.Cells(i, "AA").Value), _
            nomFichier, vbTextCompare) = 0 Then

            ObtenirCheminFichier = _
                CStr(ws.Cells(i, "AB").Value)

            Exit Function

        End If

    Next i

End Function

'========================================================
' 12. ACTUALISATION DU TABLEAU DE BORD
'========================================================

Public Sub ActualiserTableauBord()

    Dim ws As Worksheet
    Dim wsImport As Worksheet
    Dim nombreLignes As Long

    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE)

    If FeuilleExiste(FEUILLE_IMPORT) Then

        Set wsImport = ThisWorkbook.Worksheets(FEUILLE_IMPORT)

        nombreLignes = Application.WorksheetFunction.CountA( _
            wsImport.UsedRange)

        ws.Shapes("ValeurImport").TextFrame2.TextRange.Text = _
            nombreLignes & " CELLULES RENSEIGNÉES"

    Else

        ws.Shapes("ValeurImport").TextFrame2.TextRange.Text = _
            "AUCUNE IMPORTATION"

    End If

    ws.Shapes("StatutImport").TextFrame2.TextRange.Text = _
        "Dernière actualisation : " & _
        Format(Now, "dd/mm/yyyy hh:mm")

End Sub

'========================================================
' 13. VERIFICATION DES FEUILLES
'========================================================

Private Function FeuilleExiste(ByVal nom As String) As Boolean

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

        If StrComp(ws.Name, nom, vbTextCompare) = 0 Then

            FeuilleExiste = True
            Exit Function

        End If

    Next ws

End Function

'========================================================
' 14. VERIFICATION DU FORMAT
'========================================================

Private Function EstFichierExcel(ByVal chemin As String) As Boolean

    Dim extension As String
    Dim position As Long

    position = InStrRev(chemin, ".")

    If position = 0 Then Exit Function

    extension = LCase$(Mid$(chemin, position + 1))

    If extension = "xlsx" Then
        EstFichierExcel = True
    ElseIf extension = "xlsm" Then
        EstFichierExcel = True
    ElseIf extension = "xls" Then
        EstFichierExcel = True
    End If

End Function

'========================================================
' 15. NAVIGATION - TABLEAU DE BORD
'========================================================

Public Sub AllerTableauBord()

    ThisWorkbook.Worksheets(NOM_FEUILLE).Activate

    ActualiserTableauBord

End Sub

'========================================================
' 16. NAVIGATION - IMPORTATION
'========================================================

Public Sub OuvrirImportation()

    ThisWorkbook.Worksheets(NOM_FEUILLE).Activate

    ThisWorkbook.Worksheets(NOM_FEUILLE).Range( _
        CELLULE_FICHIER).Select

End Sub

'========================================================
' 17. AUTRES FONCTIONNALITES DU MENU
'========================================================

Public Sub NouveauDossier()

    MsgBox "Module NOUVEAU DOSSIER." & vbCrLf & _
        "Fonctionnalité à personnaliser selon votre structure.", _
        vbInformation

End Sub

Public Sub OuvrirConsultation()

    If FeuilleExiste(FEUILLE_IMPORT) Then

        ThisWorkbook.Worksheets(FEUILLE_IMPORT).Activate

    Else

        MsgBox "Aucune donnée importée à consulter.", vbInformation

    End If

End Sub

Public Sub OuvrirRapports()

    MsgBox "Module RAPPORTS." & vbCrLf & _
        "Espace réservé à la génération des rapports.", _
        vbInformation

End Sub

Public Sub OuvrirParametres()

    MsgBox "PARAMÈTRES DU PORTAIL" & vbCrLf & _
        "Version applicative : 1.0", vbInformation

End Sub

Public Sub OuvrirAssistance()

    MsgBox "ASSISTANCE" & vbCrLf & _
        "Portail de gestion documentaire NAVAL GROUP.", _
        vbInformation

End Sub
