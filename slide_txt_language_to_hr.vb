Sub slide_txt_language_to_hr()

    Dim slide As Slide
    Dim shape As Shape

    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                shape.TextFrame.TextRange.LanguageID = msoLanguageIDCroatian
            End If
        Next shape
    Next slide

End Sub