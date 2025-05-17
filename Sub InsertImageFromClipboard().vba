Sub InsertImageFromClipboard()
    Dim newSlide As slide
    Dim pastedShape As Shape
    Dim slideWidth As Single, slideHeight As Single
    Dim scaleFactor As Single
    
    ' Add a new blank slide
    Set newSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    
    ' Try to paste clipboard contents
    On Error Resume Next
    newSlide.Shapes.Paste
    If Err.Number <> 0 Then
        MsgBox "Nothing pasteable found on the clipboard. Please copy an image first.", vbExclamation, "Paste Failed"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get the most recently pasted shape
    Set pastedShape = newSlide.Shapes(newSlide.Shapes.Count)
    
    ' Get slide dimensions
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
    
    ' Scale to fit while maintaining aspect ratio
    If pastedShape.Width / pastedShape.Height > slideWidth / slideHeight Then
        scaleFactor = slideWidth / pastedShape.Width
    Else
        scaleFactor = slideHeight / pastedShape.Height
    End If

    With pastedShape
        .LockAspectRatio = msoTrue
        .Width = .Width * scaleFactor
        .Height = .Height * scaleFactor
        .Left = (slideWidth - .Width) / 2
        .Top = (slideHeight - .Height) / 2
    End With
End Sub

