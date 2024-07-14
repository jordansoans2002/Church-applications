Sub CreatePresentationFromTextFiles(fileNames As Variant, titleFontName As String, titleFontSize As Integer, lyricsFontName As String, lyricsFontSize As Integer)
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim titleSlide As Object
    Dim i As Integer
    Dim fileName As String
    Dim lyricsText As String
    Dim slides() As String
    Dim slideIndex As Integer
    
    ' Create PowerPoint application object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Loop through each filename passed as input
    For i = LBound(fileNames) To UBound(fileNames)
        fileName = fileNames(i)
        
        ' Read the entire text from the file
        Open fileName For Input As #1
        lyricsText = Input$(LOF(1), 1)
        Close #1
        
        ' Split text into slides based on double newline
        slides = Split(lyricsText, vbCrLf & vbCrLf)
        
        ' Create title slide for the song
        Set titleSlide = pptPres.slides.Add(pptPres.slides.Count + 1, 1) ' ppLayoutTitle = 1
        
        ' Set title slide text and format
        titleSlide.Shapes.Title.TextFrame.TextRange.Text = Left(GetFileNameFromPath(fileName), Len(GetFileNameFromPath(fileName)) - 4) ' Remove ".txt" extension from title
        With titleSlide.Shapes.Title.TextFrame.TextRange.Font
            .Name = titleFontName
            .Size = titleFontSize
        End With
        
        ' Center title text horizontally and vertically
        With titleSlide.Shapes.Title.TextFrame.TextRange.ParagraphFormat
            .Alignment = 2 ' ppAlignCenter = 2
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        
        ' Loop through each slide text and add to presentation
        For slideIndex = LBound(slides) To UBound(slides)
            ' Create a new slide
            Set pptSlide = pptPres.slides.Add(pptPres.slides.Count + 1, 2) ' ppLayoutText = 2
            
            ' Set lyrics slide text and format
            pptSlide.Shapes(1).TextFrame.TextRange.Text = slides(slideIndex)
            With pptSlide.Shapes(1).TextFrame.TextRange.Font
                .Name = lyricsFontName
                .Size = lyricsFontSize
            End With
            
            ' Center lyrics text horizontally and vertically
            With pptSlide.Shapes(1).TextFrame.TextRange.ParagraphFormat
                .Alignment = 2 ' ppAlignCenter = 2
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
        Next slideIndex
    Next i
    
    ' Clean up
    Set pptSlide = Nothing
    Set titleSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
End Sub

Function GetFileNameFromPath(filePath As String) As String
    GetFileNameFromPath = Mid(filePath, InStrRev(filePath, "\") + 1)
End Function

Sub RunPresentationCreation()
    Dim fileNames As Variant
    Dim titleFontName As String
    Dim titleFontSize As Integer
    Dim lyricsFontName As String
    Dim lyricsFontSize As Integer
    
    ' Define the list of .txt files (replace with your file paths)
    fileNames = Array("C:\Users\admin\Desktop\Lyrics.txt")
    
    ' Define font properties
    titleFontName = "Arial"
    titleFontSize = 24
    lyricsFontName = "Calibri"
    lyricsFontSize = 48
    
    ' Call the main function to create the presentation
    CreatePresentationFromTextFiles fileNames, titleFontName, titleFontSize, lyricsFontName, lyricsFontSize
End Sub

