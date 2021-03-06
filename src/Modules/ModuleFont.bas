Sub CopyFontAttributeToClipboard()

    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object

    Dim fontFace As String
    Dim fontSize As Single
    Dim color As MsoRGBType
    Dim hexColor As String
    Dim bold As Boolean
    Dim italic As Boolean
    Dim charSpacing As Single
    Dim align As String
    Dim lineSpacing As Single
    Dim paraSpaceBefore As Single
    Dim paraSpaceAfter As Single
    'Dim text As String
    Dim valign As String
    Dim marginLeft As Single
    Dim marginRight As Single
    Dim marginBottom As Single
    Dim marginTop As Single
    Dim autoSize As PpAutoSize

    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    Dim PlaceHolderTextRange As TextRange
    Set PlaceHolderTextRange = SlidePlaceHolder.TextFrame.TextRange

    Set myDocument = Application.ActiveWindow

    If Not myDocument.Selection.Type = ppSelectionText Then
        SlidePlaceHolder.Delete
        MsgBox "No text selected."
    Else
        Set FontObj = myDocument.Selection.TextRange.Font
        Set ParagraphObj = myDocument.Selection.TextRange.ParagraphFormat
        Set TextParent = myDocument.Selection.TextRange.Parent

        fontFace = FontObj.Name
        fontSize = FontObj.Size
        color = FontObj.Color.RGB
        bold = FontObj.Bold
        italic = FontObj.Italic
        #if Mac then
          'Run on Mac
           charSpacing = FontObj.Spacing
        #else
          'Run on PC
          'TODO It seems all ppt versions on PC don't support Spacing
        #endif

        align = ParagraphObj.Alignment
        lineSpacing = ParagraphObj.SpaceWithin
        paraSpaceBefore = ParagraphObj.SpaceBefore
        paraSpaceAfter = ParagraphObj.SpaceAfter

        valign = TextParent.VerticalAnchor
        marginLeft = TextParent.marginLeft
        marginRight = TextParent.marginRight
        marginBottom = TextParent.marginBottom
        marginTop = TextParent.MarginTop

        autoSize = TextParent.AutoSize

        'text = myDocument.Selection.TextRange.Text

        'font attribute
        PlaceHolderTextRange.Characters(0).InsertAfter "fontFace: '" & fontFace & "'," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "fontSize: " & Round(fontSize, 3) & "," & Chr(13)
        If Not color = 0 Then
            hexColor = Right("000000" & Hex(color), 6)
            hexColor = "#" & Right(hexColor, 2) & Mid(hexColor, 3, 2) & Left(hexColor, 2)
            PlaceHolderTextRange.Characters(0).InsertAfter "color: '" & hexColor & "'," & Chr(13)
        End If
        If bold Then
            PlaceHolderTextRange.Characters(0).InsertAfter "bold: true," & Chr(13)
        End If
        If italic Then
            PlaceHolderTextRange.Characters(0).InsertAfter "italic: true," & Chr(13)
        End If
        'unit is points
        if Not charSpacing = 0 Then
            PlaceHolderTextRange.Characters(0).InsertAfter "charSpacing: " & Round(charSpacing, 3) & "," & Chr(13)
        End If

        'paragraph attribute
        If align = ppAlignLeft Then
            PlaceHolderTextRange.Characters(0).InsertAfter "align: 'left'," & Chr(13)
        ElseIf align = ppAlignCenter Then
            PlaceHolderTextRange.Characters(0).InsertAfter "align: 'center'," & Chr(13)
        ElseIf align = ppAlignRight Then
            PlaceHolderTextRange.Characters(0).InsertAfter "align: 'right'," & Chr(13)
        End If
        'unit is lineSpacing percent
        If ParagraphObj.LineRuleWithin Then
            If Not lineSpacing = 1 Then
                PlaceHolderTextRange.Characters(0).InsertAfter "lineSpacingMultiple: " & Round(lineSpacing, 3) & "," & Chr(13)
            End If
        'unit is points
        ElseIf Not lineSpacing = 0 Then
            PlaceHolderTextRange.Characters(0).InsertAfter "lineSpacing: " & Round(lineSpacing, 3) & "," & Chr(13)
        End If
        'unit is points
        if Not paraSpaceBefore = 0 Then
            PlaceHolderTextRange.Characters(0).InsertAfter "paraSpaceBefore: " & Round(paraSpaceBefore, 3) & "," & Chr(13)
        End If
        if Not paraSpaceAfter = 0 Then
            PlaceHolderTextRange.Characters(0).InsertAfter "paraSpaceAfter: " & Round(paraSpaceAfter, 3) & "," & Chr(13)
        End If

        'other attribute
        If valign = msoAnchorTop Then
            PlaceHolderTextRange.Characters(0).InsertAfter "valign: 'top'," & Chr(13)
        ElseIf valign = msoAnchorMiddle Then
            PlaceHolderTextRange.Characters(0).InsertAfter "valign: 'middle'," & Chr(13)
        ElseIf valign = msoAnchorBottom Then
            PlaceHolderTextRange.Characters(0).InsertAfter "valign: 'bottom'," & Chr(13)
        End If

        'fit attribute
        If autoSize = ppAutoSizeNone Then
            PlaceHolderTextRange.Characters(0).InsertAfter "fit: 'none'," & Chr(13)
        ElseIf autoSize = ppAutoSizeShapeToFitText Then
            PlaceHolderTextRange.Characters(0).InsertAfter "fit: 'resize'," & Chr(13)
        ElseIf autoSize = ppAutoSizeMixed Then
            PlaceHolderTextRange.Characters(0).InsertAfter "fit: 'shrink'," & Chr(13)
        End If

        'unit is points
        PlaceHolderTextRange.Characters(0).InsertAfter "margin: [" & Round(marginLeft, 3) & ", " & Round(marginRight, 3) & ", " & Round(marginBottom, 3) & ", " & Round(marginTop, 3) & "]," & Chr(13)

        SlidePlaceHolder.TextFrame.TextRange.Copy
        SlidePlaceHolder.Delete

    End If

End Sub
