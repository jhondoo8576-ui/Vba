
Option Explicit

' ==========================================
' CONFIGURATION VISUELLE
' ==========================================

Private Const BG As Long = 15790320
Private Const NAVY As Long = 1050624
Private Const BLUE As Long = 1518549
Private Const WHITE As Long = 16777215
Private Const GREY As Long = 15132390
Private Const ACCENT As Long = 4899327

' ==========================================
' CONSTRUCTION DE L'INTERFACE
' ==========================================

Public Sub GenererInterface()

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = ThisWorkbook.Worksheets("TEST")

    Application.ScreenUpdating = False

    On Error GoTo GestionErreur

    ws.Activate

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False

    ' Nettoyage de l'interface
    ws.Cells.Clear

    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ws.Cells.Interior.Color = BG
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ws.Rows.RowHeight = 25
    ws.Columns.ColumnWidth = 4

    ActiveWindow.Zoom = 85

    ' ======================================
    ' HEADER
    ' ======================================

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 1050, 85)

    shp.Fill.ForeColor.RGB = NAVY
    shp.Line.Visible = msoFalse

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 25, 17, 52, 52)

    shp.Fill.ForeColor.RGB = WHITE
    shp.Line.Visible = msoFalse

    With shp.TextFrame2
        .TextRange.Text = "NG"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 18
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = NAVY
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 95, 19, 500, 30)

    With shp.TextFrame2
        .TextRange.Text = "NAVAL GROUP"
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 22
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = WHITE
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 97, 50, 500, 20)

    With shp.TextFrame2
        .TextRange.Text = "ESPACE DE TRAVAIL | PORTAIL DE GESTION"
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = RGB(190, 205, 220)
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' ======================================
    ' TITRE CENTRAL
    ' ======================================

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 115, 900, 45)

    With shp.TextFrame2
        .TextRange.Text = "Bienvenue dans votre espace"
        .TextRange.Font.Size = 25
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = NAVY
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 52, 157, 900, 25)

    With shp.TextFrame2
        .TextRange.Text = "Sélectionnez une fonctionnalité pour commencer."
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 110, 125)
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' ======================================
    ' MENU PRINCIPAL
    ' ======================================

    Call CreerBouton(ws, "TABLEAU DE BORD", "Accueil", 65, 225, BLUE)
    Call CreerBouton(ws, "NOUVEAU DOSSIER", "NouveauDossier", 355, 225, NAVY)
    Call CreerBouton(ws, "CONSULTATION", "Consultation", 645, 225, NAVY)

    Call CreerBouton(ws, "RAPPORTS", "Rapports", 65, 335, NAVY)
    Call CreerBouton(ws, "PARAMÈTRES", "Parametres", 355, 335, NAVY)
    Call CreerBouton(ws, "ASSISTANCE", "Assistance", 645, 335, NAVY)

    ' ======================================
    ' ZONE DE SELECTION
    ' ======================================

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 65, 445, 870, 150)

    shp.Fill.ForeColor.RGB = WHITE
    shp.Line.ForeColor.RGB = GREY
    shp.Line.Weight = 1

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 90, 455, 800, 25)

    With shp.TextFrame2
        .TextRange.Text = "SÉLECTION RAPIDE"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = BLUE
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' Libellé service
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 95, 490, 150, 20)

    With shp.TextFrame2
        .TextRange.Text = "Choisir un service :"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = NAVY
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' Liste déroulante service
    Call CreerListeDeroulante(ws, "cmbService", _
        95, 515, 250, 28, _
        Array("Production", "Bureau d'études", "Maintenance", _
              "Qualité", "Logistique", "Administration"))

    ' Libellé statut
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 390, 490, 150, 20)

    With shp.TextFrame2
        .TextRange.Text = "Choisir un statut :"
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = NAVY
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' Liste déroulante statut
    Call CreerListeDeroulante(ws, "cmbStatut", _
        390, 515, 220, 28, _
        Array("En attente", "En cours", "Terminé", "Annulé"))

    ' Bouton validation
    Call CreerBouton(ws, "VALIDER", "ValiderSelection", _
                     680, 510, BLUE)

    ' Réduction de la taille du bouton
    ws.Shapes(ws.Shapes.Count).Width = 180
    ws.Shapes(ws.Shapes.Count).Height = 38

    ' ======================================
    ' FOOTER
    ' ======================================

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 630, 1050, 35)

    shp.Fill.ForeColor.RGB = NAVY
    shp.Line.Visible = msoFalse

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 25, 637, 1000, 20)

    With shp.TextFrame2
        .TextRange.Text = "NAVAL GROUP | ENVIRONNEMENT DE TRAVAIL | VERSION 2.0"
        .TextRange.Font.Size = 8
        .TextRange.Font.Fill.ForeColor.RGB = WHITE
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' ======================================
    ' FINALISATION
    ' ======================================

    ws.Activate

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 85

    Application.ScreenUpdating = True

    Exit Sub

GestionErreur:

    Application.ScreenUpdating = True

    MsgBox "Erreur lors de la création de l'interface : " _
           & Err.Description, vbCritical

End Sub

' ==========================================
' CREATION DES BOUTONS
' ==========================================

Private Sub CreerBouton(ByVal ws As Worksheet, _
                        ByVal titre As String, _
                        ByVal macro As String, _
                        ByVal x As Double, _
                        ByVal y As Double, _
                        ByVal couleur As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, 250, 75)

    shp.Fill.ForeColor.RGB = couleur
    shp.Line.Visible = msoFalse

    With shp.TextFrame2
        .TextRange.Text = titre
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = WHITE
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    shp.OnAction = macro

End Sub

' ==========================================
' CREATION DES LISTES DEROULANTES
' ==========================================

Private Sub CreerListeDeroulante(ByVal ws As Worksheet, _
                                 ByVal nom As String, _
                                 ByVal x As Double, _
                                 ByVal y As Double, _
                                 ByVal largeur As Double, _
                                 ByVal hauteur As Double, _
                                 ByVal elements As Variant)

    Dim obj As OLEObject
    Dim i As Long

    Set obj = ws.OLEObjects.Add( _
        ClassType:="Forms.ComboBox.1", _
        Left:=x, _
        Top:=y, _
        Width:=largeur, _
        Height:=hauteur)

    With obj

        .Name = nom
        .Placement = xlFreeFloating

        With .Object

            .Clear

            For i = LBound(elements) To UBound(elements)
                .AddItem elements(i)
            Next i

            .Style = 2
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .BackColor = WHITE
            .ForeColor = NAVY
            .MatchRequired = True

        End With

    End With

End Sub

' ==========================================
' VALIDATION DES LISTES
' ==========================================

Public Sub ValiderSelection()

    Dim ws As Worksheet
    Dim service As String
    Dim statut As String

    Set ws = ThisWorkbook.Worksheets("TEST")

    service = ws.OLEObjects("cmbService").Object.Value
    statut = ws.OLEObjects("cmbStatut").Object.Value

    If service = "" Or statut = "" Then

        MsgBox "Veuillez sélectionner un service et un statut.", _
               vbExclamation, "Sélection incomplète"

        Exit Sub

    End If

    MsgBox "Sélection enregistrée :" & vbCrLf & vbCrLf & _
           "Service : " & service & vbCrLf & _
           "Statut : " & statut, _
           vbInformation, "Validation"

End Sub

' ==========================================
' ACTIONS DES BOUTONS
' ==========================================

Public Sub Accueil()

    MsgBox "Bienvenue dans votre tableau de bord.", vbInformation

End Sub

Public Sub NouveauDossier()

    MsgBox "Module Nouveau Dossier", vbInformation

End Sub

Public Sub Consultation()

    MsgBox "Module Consultation", vbInformation

End Sub

Public Sub Rapports()

    MsgBox "Module Rapports", vbInformation

End Sub

Public Sub Parametres()

    MsgBox "Module Paramètres", vbInformation

End Sub

Public Sub Assistance()

    MsgBox "Module Assistance", vbInformation

End Sub
