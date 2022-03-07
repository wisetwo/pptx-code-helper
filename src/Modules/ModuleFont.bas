Sub CopyFontAttributeToClipboard()

    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object

    Dim fontFace As String
    Dim fontSize As Long
    Dim color As Long
    Dim hexColor As String
    Dim bold As Boolean
    Dim italic As Boolean
    Dim align As String
    'Dim valign As String
    'Dim lineSpacing As Long
    'Dim text As String

    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    Dim PlaceHolderTextRange As TextRange
    Set PlaceHolderTextRange = SlidePlaceHolder.TextFrame.TextRange

    Set myDocument = Application.ActiveWindow

    If Not myDocument.Selection.Type = ppSelectionText Then
        SlidePlaceHolder.Delete
        MsgBox "No text selected."
    Else
        fontFace = myDocument.Selection.TextRange.Font.Name
        fontSize = myDocument.Selection.TextRange.Font.Size
        color = myDocument.Selection.TextRange.Font.Color.RGB
        bold = myDocument.Selection.TextRange.Font.Bold
        italic = myDocument.Selection.TextRange.Font.Italic
        align = myDocument.Selection.TextRange.ParagraphFormat.Alignment
        'valign = ?
        ' lineSpacing = myDocument.Selection.TextRange.ParagraphFormat.SpaceWithin
        'text = myDocument.Selection.TextRange.Text

        PlaceHolderTextRange.Characters(0).InsertAfter Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "fontFace: """ & fontFace & """," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "fontSize: " & Round(fontSize, 2) & "," & Chr(13)
        If Not color = 0 Then
            hexColor = Right("000000" & Hex(color), 6)
            hexColor = "#" & Right(hexColor, 2) & Mid(hexColor, 3, 2) & Left(hexColor, 2)
            PlaceHolderTextRange.Characters(0).InsertAfter "color: """ & hexColor & """" & Chr(13)
        End If
        If bold Then
            PlaceHolderTextRange.Characters(0).InsertAfter "bold: true," & Chr(13)
        End If
        If italic Then
            PlaceHolderTextRange.Characters(0).InsertAfter "italic: true," & Chr(13)
        End If
        If align = ppAlignCenter Then
            PlaceHolderTextRange.Characters(0).InsertAfter "align: ""center""," & Chr(13)
        ElseIf align = ppAlignRight Then
            PlaceHolderTextRange.Characters(0).InsertAfter "align: ""right""," & Chr(13)
        End If
        ' PlaceHolderTextRange.Characters(0).InsertAfter "lineSpacing: " & Round(lineSpacing, 2) & "," & Chr(13)

        SlidePlaceHolder.TextFrame.TextRange.Copy
        SlidePlaceHolder.Delete

    End If

End Sub
