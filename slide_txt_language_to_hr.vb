Sub slide_txt_language_to_hr()
    ' Declare variables 'slide' and 'shape' that represent a slide in presentation and shape on a slide.
    Dim slide As Slide
    Dim shape As Shape
    ' Iterate through all slides in the active presentation.
    For Each slide In ActivePresentation.Slides
        ' Iterate through all shapes on each slide.
        For Each shape In slide.Shapes
            ' Check if the current shape has a text frame.
            If shape.HasTextFrame Then
                ' Set the language of the text in the text range of the text frame to Croatian.
                shape.TextFrame.TextRange.LanguageID = msoLanguageIDCroatian
            End If
        Next shape
    Next slide
End Sub