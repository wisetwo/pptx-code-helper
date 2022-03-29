Sub CopyShapeAttributeToClipboard()

    Dim PresentationSlide As PowerPoint.Slide
    Dim SlidePlaceHolder As PowerPoint.Shape
    Dim ClipboardObject As Object
    Dim PosX As Long
    Dim PosY As Long
    Dim Width As Long
    Dim Height As Long

    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)
    Dim PlaceHolderTextRange As TextRange
    Set PlaceHolderTextRange = SlidePlaceHolder.TextFrame.TextRange

    Set myDocument = Application.ActiveWindow

    If Not myDocument.Selection.Type = ppSelectionShapes Then
        SlidePlaceHolder.Delete
        MsgBox "No shapes selected."
    Else
        Set ShapeObj = myDocument.Selection.ShapeRange
        PosX = ShapeObj.Left
        PosY = ShapeObj.Top
        Width = ShapeObj.Width
        Height = ShapeObj.Height
        PlaceHolderTextRange.Characters(0).InsertAfter "x: " & Round(PosX / 72, 3) & "," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "y: " & Round(PosY / 72, 3) & "," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "w: " & Round(Width / 72, 3) & "," & Chr(13)
        PlaceHolderTextRange.Characters(0).InsertAfter "h: " & Round(Height / 72, 3) & "," & Chr(13)

        SlidePlaceHolder.TextFrame.TextRange.Copy
        SlidePlaceHolder.Delete

    End If

End Sub
