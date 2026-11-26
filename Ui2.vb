
Option Explicit

Public Sub CreerDropDown( _
    ByVal ws As Worksheet, _
    ByVal nom As String, _
    ByVal elements As Variant, _
    ByVal x As Double, _
    ByVal y As Double, _
    Optional ByVal largeur As Double = 250, _
    Optional ByVal hauteur As Double = 40)

    Dim shp As Shape
    Dim i As Long
    Dim nb As Long

    nb = UBound(elements) - LBound(elements) + 1

    ' Supprimer l'ancien composant
    On Error Resume Next

    ws.Shapes("DD_" & nom).Delete

    For i = 1 To nb
        ws.Shapes("DD_ITEM_" & nom & "_" & i).Delete
    Next i

    On Error GoTo 0

    ' ==================================
    ' CHAMP PRINCIPAL
    ' ==================================

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        x, y, largeur, hauteur)

    With shp

        .Name = "DD_" & nom

        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(210, 220, 230)
        .Line.Weight = 1

        With .TextFrame2

            .TextRange.Text = "Sélectionner une option..."
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 10
            .TextRange.Font.Fill.ForeColor.RGB = RGB(16, 32, 64)

            .MarginLeft = 12
            .MarginRight = 30
            .VerticalAnchor = msoAnchorMiddle

        End With

        .OnAction = "OuvrirDropDown"

        .AlternativeText = nom

    End With

    ' ==================================
    ' FLECHE
    ' ==================================

    Set shp = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        x + largeur - 32, y + 5, 25, hauteur - 10)

    With shp

        .Name = "DD_ARROW_" & nom

        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse

        With .TextFrame2

            .TextRange.Text = ChrW(&H25BE)
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 17
            .TextRange.Font.Fill.ForeColor.RGB = RGB(23, 82, 145)

            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle

        End With

        .OnAction = "OuvrirDropDown"
        .AlternativeText = nom

    End With

    ' ==================================
    ' ELEMENTS DE LA LISTE
    ' ==================================

    For i = LBound(elements) To UBound(elements)

        Set shp = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            x, _
            y + hauteur + ((i - LBound(elements)) * 34), _
            largeur, 32)

        With shp

            .Name = "DD_ITEM_" & nom & "_" & _
                    (i - LBound(elements) + 1)

            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Line.ForeColor.RGB = RGB(220, 225, 232)
            .Line.Weight = 0.5

            With .TextFrame2

                .TextRange.Text = CStr(elements(i))
                .TextRange.Font.Name = "Segoe UI"
                .TextRange.Font.Size = 10
                .TextRange.Font.Fill.ForeColor.RGB = RGB(16, 32, 64)

                .MarginLeft = 12
                .VerticalAnchor = msoAnchorMiddle

            End With

            .AlternativeText = nom
            .OnAction = "ChoisirDropDown"
            .Visible = msoFalse

        End With

    Next i

    ' Mémoriser le nombre d'options
    ws.Shapes("DD_" & nom).Tags.Add "Nombre", CStr(nb)

End Sub

' ==========================================
' OUVRIR / FERMER LA LISTE
' ==========================================

Public Sub OuvrirDropDown()

    Dim ws As Worksheet
    Dim nom As String
    Dim nb As Long
    Dim i As Long
    Dim visible As Boolean

    Set ws = ActiveSheet

    nom = ws.Shapes(Application.Caller).AlternativeText

    nb = CLng(ws.Shapes("DD_" & nom).Tags("Nombre"))

    visible = ws.Shapes("DD_ITEM_" & nom & "_1").Visible

    For i = 1 To nb

        ws.Shapes("DD_ITEM_" & nom & "_" & i).Visible = Not visible

    Next i

End Sub

' ==========================================
' SELECTIONNER UNE OPTION
' ==========================================

Public Sub ChoisirDropDown()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim nom As String
    Dim nb As Long
    Dim i As Long
    Dim valeur As String

    Set ws = ActiveSheet

    Set shp = ws.Shapes(Application.Caller)

    nom = shp.AlternativeText
    valeur = shp.TextFrame2.TextRange.Text

    nb = CLng(ws.Shapes("DD_" & nom).Tags("Nombre"))

    ' Mettre à jour la valeur sélectionnée
    ws.Shapes("DD_" & nom).TextFrame2.TextRange.Text = valeur

    ' Fermer la liste
    For i = 1 To nb

        ws.Shapes("DD_ITEM_" & nom & "_" & i).Visible = msoFalse

    Next i

End Sub
