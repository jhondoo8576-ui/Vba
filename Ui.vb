
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
    Dim i As Integer

    Set ws = ThisWorkbook.Worksheets("TEST")

    Application.ScreenUpdating = False

    ' Configuration de la feuille
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False

    ws.Cells.Clear
    ws.Cells.Interior.Color = BG
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ws.Rows.RowHeight = 25
    ws.Columns.ColumnWidth = 4

    ' Masquer les colonnes et lignes inutiles
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

    shp.TextFrame2.TextRange.Text = "NG"
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 18
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = NAVY
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 95, 19, 500, 30)

    shp.TextFrame2.TextRange.Text = "NAVAL GROUP"
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 22
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = WHITE
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 97, 50, 500, 20)

    shp.TextFrame2.TextRange.Text = "ESPACE DE TRAVAIL | PORTAIL DE GESTION"
    shp.TextFrame2.TextRange.Font.Size = 9
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(190, 205, 220)
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' ======================================
    ' TITRE CENTRAL
    ' ======================================

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 115, 900, 45)

    shp.TextFrame2.TextRange.Text = "Bienvenue dans votre espace"
    shp.TextFrame2.TextRange.Font.Size = 25
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = NAVY
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 52, 157, 900, 25)

    shp.TextFrame2.TextRange.Text = "Sélectionnez une fonctionnalité pour commencer."
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(100, 110, 125)
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
    ' ZONE CENTRALE
    ' ======================================

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 65, 465, 870, 100)

    shp.Fill.ForeColor.RGB = WHITE
    shp.Line.ForeColor.RGB = GREY
    shp.Line.Weight = 1

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 90, 480, 800, 25)

    shp.TextFrame2.TextRange.Text = "INFORMATION"
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BLUE
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 90, 510, 800, 35)

    shp.TextFrame2.TextRange.Text = "Bienvenue sur votre interface de gestion. Utilisez le menu pour accéder aux différents modules."
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = NAVY
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse

    ' ======================================
    ' FOOTER
    ' ======================================

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 610, 1050, 35)

    shp.Fill.ForeColor.RGB = NAVY
    shp.Line.Visible = msoFalse

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 25, 617, 1000, 20)

    shp.TextFrame2.TextRange.Text = "NAVAL GROUP  |  ENVIRONNEMENT DE TRAVAIL  |  VERSION 1.0"
    shp.TextFrame2.TextRange.Font.Size = 8
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = WHITE
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
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

    shp.TextFrame2.TextRange.Text = titre
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = WHITE

    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    shp.OnAction = macro

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
