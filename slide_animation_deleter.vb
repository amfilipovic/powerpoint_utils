Sub slide_animation_deleter()

    Dim slide As Slide
    Dim shape As Shape

    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            shape.AnimationSettings.Animate = FALSE
        Next shape
    Next slide

End Sub