Sub CreatePresentationFromTextFiles(songNames As Variant, titleFontName As String, titleFontSize As Integer, lyricsFontName As String, lyricsFontSize As Integer)
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim titleSlide As Object
    Dim i As Integer
    Dim songName As String
    Dim lyricsText As String
    Dim slides() As String
    Dim slideIndex As Integer
    Dim fs As Object
    Dim ts As Object
    
    ' Create PowerPoint application object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Loop through each songName passed as input
    For i = LBound(songNames) To UBound(songNames)
        songName = songNames(i)
        filePath = "C:\Users\admin\Desktop\" + songName
        
        ' Check if the file exists
        If Dir(filePath + "_english.txt") <> "" Then
            ' Read the entire text from the file
            Open filePath + "_english.txt" For Input As #1
            eng_lyrics = Input$(LOF(1), 1)
            Close #1
        Else
            eng_lyrics = ""
        End If
        
         ' Check if the file exists
        If Dir(filePath + "_hindi.txt") <> "" Then
             ' Read the entire text from the file
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set ts = fs.OpenTextFile(filePath + "_hindi.txt", 1, False, -1) ' Open file for reading with specified encoding
            hin_lyrics = ts.ReadAll
            ts.Close
            Set ts = Nothing
            Set fs = Nothing
        Else
            hin_lyrics = ""
        End If
    
        
        ' Find the position of the end of the first line
        posFirstLineEnd = InStr(eng_lyrics, vbCrLf)
    
        ' Extract the first line
        If posFirstLineEnd > 0 Then
            citation = Left(eng_lyrics, posFirstLineEnd - 1)
        End If
        ' Extract the remaining text after removing the first line
        eng_lyrics = Mid(eng_lyrics, posFirstLineEnd + Len(vbCrLf) * 2)
        
        ' Split text into slides based on double newline
        slides = Split(eng_lyrics, vbCrLf & vbCrLf)
        hin_slides = Split(hin_lyrics, vbCrLf & vbCrLf)
        
        Debug.Print "eng slides " & UBound(slides) - LBound(slides) + 1
        Debug.Print "hindi slides " & UBound(hin_slides) - LBound(hin_slides) + 1
        
        ' Create title slide for the song
        Set titleSlide = pptPres.slides.Add(pptPres.slides.Count + 1, 1) ' ppLayoutTitle = 1
        
        ' Set title slide text and format
        titleSlide.Shapes.Title.textFrame.TextRange.Text = songName
        With titleSlide.Shapes.Title.textFrame.TextRange.Font
            .Name = titleFontName
            .Size = titleFontSize
        End With
        
        ' Center title text horizontally and vertically
        With titleSlide.Shapes.Title.textFrame.TextRange.ParagraphFormat
            .Alignment = 2 ' ppAlignCenter = 2
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        
        Set footerShape = titleSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 18, titleSlide.Master.Height - 36, titleSlide.Master.Width, titleSlide.Master.Height - 24)
        footerShape.textFrame.TextRange.Text = citation
        ' Set lyrics slide text in the left text box
        With footerShape.textFrame.TextRange.Font
            .Name = lyricsFontName
            .Size = 18
        End With
        ' Center the lyrics horizontally and vertically
        With footerShape.textFrame.TextRange.ParagraphFormat
            .Alignment = 1
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        
        ' Loop through each slide text and add to presentation
        For slideIndex = LBound(slides) To UBound(slides)
            ' Create a new slide with two-column text layout
            Set pptSlide = pptPres.slides.Add(pptPres.slides.Count + 1, ppLayoutBlank) ' ppLayoutTwoColumnText = 3
            SlideWidth = pptSlide.Master.Width
            slideHeight = pptSlide.Master.Height
            
            ' Calculate dimensions for textboxes
            textboxWidth = SlideWidth / 2
            textboxHeight = slideHeight
    
            ' Add first textbox
            Set pptTextbox1 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 18, 32, textboxWidth - 18, textboxHeight - 36)
            pptTextbox1.textFrame.TextRange.Text = slides(slideIndex) ' Mid(slides(slideIndex), 0)
            ' Set lyrics slide text in the left text box
            With pptTextbox1.textFrame.TextRange.Font
                .Name = "Times New Roman"
                .Size = lyricsFontSize
            End With
            ' Center the lyrics horizontally and vertically
            With pptTextbox1.textFrame.TextRange.ParagraphFormat
                .Alignment = 2
                .SpaceBefore = 0
                .SpaceAfter = 0
                ' .LineRuleWithin = 1 ' ppLineSpaceSingle = 1 (Set single line spacing)
                ' .SpaceWithin = 1.5 ' Adjust line spacing factor here (1.5 = 1.5 times the default line spacing)
            End With
            
    
            ' Add second textbox
            Set pptTextbox2 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, textboxWidth + 18, 32, textboxWidth - 36, textboxHeight - 36)
            pptTextbox2.textFrame.TextRange.Text = hin_slides(slideIndex) 'Mid(hin_slides(slideIndex), 0)
             ' Set lyrics slide text in the left text box
            With pptTextbox2.textFrame.TextRange.Font
                .Name = "Mangal"
                .Size = lyricsFontSize - 2
            End With
            ' Center the lyrics horizontally and vertically
            With pptTextbox2.textFrame.TextRange.ParagraphFormat
                .Alignment = 2
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineRuleWithin = 1 ' ppLineSpaceSingle = 1 (Set single line spacing)
                .SpaceWithin = 1 ' Adjust line spacing factor here (1.5 = 1.5 times the default line spacing)
            End With
            
            Set footerShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 18, textboxHeight - 36, SlideWidth, textboxHeight - 24)
            footerShape.textFrame.TextRange.Text = citation
             ' Set lyrics slide text in the left text box
            With footerShape.textFrame.TextRange.Font
                .Name = lyricsFontName
                .Size = 18
            End With
            ' Center the lyrics horizontally and vertically
            With footerShape.textFrame.TextRange.ParagraphFormat
                .Alignment = 1
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


Sub RunPresentationCreation()
    Dim songNames As Variant
    Dim titleFontName As String
    Dim titleFontSize As Integer
    Dim lyricsFontName As String
    Dim lyricsFontSize As Integer
    
    ' Define the list of .txt files (replace with your file paths)
    songNames = Array("Lyrics")
    
    ' Define font properties
    titleFontName = "Arial"
    titleFontSize = 60
    lyricsFontName = "Times New Roman"
    lyricsFontSize = 43
    
    ' Call the main function to create the presentation
    CreatePresentationFromTextFiles songNames, titleFontName, titleFontSize, lyricsFontName, lyricsFontSize
End Sub

