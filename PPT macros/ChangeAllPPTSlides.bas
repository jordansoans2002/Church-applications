Dim keepMonitoring As Boolean

Sub NextSlideInAllPresentations()
    Dim pres As Presentation
    Dim slideIndex As Integer
    
    ' Loop through all open presentations
    For Each pres In Application.Presentations
        ' Check if a slide show is running
        If Not pres.SlideShowWindow Is Nothing Then
            ' Get the current slide index
            slideIndex = pres.SlideShowWindow.View.CurrentShowPosition
            
            ' Go to the next slide if it exists
            If slideIndex < pres.slides.Count Then
                pres.SlideShowWindow.View.GotoSlide slideIndex + 1
            End If
        End If
    Next pres
End Sub

