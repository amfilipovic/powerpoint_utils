Sub slide_animation_deleter()
    ' Declare variables 'slide' and 'shape' that represent a slide in presentation and shape on a slide.
    Dim slide As Slide
    Dim shape As Shape
    ' Iterate through all slides in the active presentation.
    For Each slide In ActivePresentation.Slides
        ' Iterate through all shapes on each slide.
        For Each shape In slide.Shapes
            ' Disable animation for the current shape by setting its 'Animate' property to 'FALSE'.
            shape.AnimationSettings.Animate = FALSE
        Next shape
    Next slide
End Sub