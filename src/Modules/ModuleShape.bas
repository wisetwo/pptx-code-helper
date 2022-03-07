Sub CopyShapeAttributeToClipboard()

    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object
    Dim TopToCopy As Long
    Dim LeftToCopy As Long
    Dim WidthToCopy As Long
    Dim HeightToCopy As Long

    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    Dim PlaceHolderTextRange As TextRange
    Set PlaceHolderTextRange = SlidePlaceHolder.TextFrame.TextRange

    Set myDocument = Application.ActiveWindow

    If Not myDocument.Selection.Type = ppSelectionShapes Then
        SlidePlaceHolder.Delete
        MsgBox "No shapes selected."
    Else
        TopToCopy = myDocument.Selection.ShapeRange(1).Top
        LeftToCopy = myDocument.Selection.ShapeRange(1).Left
        WidthToCopy = myDocument.Selection.ShapeRange(1).Width
        HeightToCopy = myDocument.Selection.ShapeRange(1).Height
        PlaceHolderTextRange.Characters(0).InsertAfter Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "x: " & Round(LeftToCopy / 72, 2) & "," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "y: " & Round(TopToCopy / 72, 2) & "," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "w: " & Round(WidthToCopy / 72, 2) & "," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "h: " & Round(HeightToCopy / 72, 2) & "," & Chr(13)

        SlidePlaceHolder.TextFrame.TextRange.Copy
        SlidePlaceHolder.Delete

    End If

End Sub
