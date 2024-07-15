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
            MsgBox "English Lyrics were not found for " & songName & ". Please ensure lyrics for the song are present in the songs folder and songname entered is correct and matches the lyrics file in the folder."
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
            MsgBox "Hindi Lyrics were not found for " & songName & ". Please ensure lyrics for the song are present in the songs folder and song name entered is correct and matches the lyrics file in the folder."
            hin_lyrics = ""
        End If
    
        
        ' Find the position of the end of the first line
        posFirstLineEnd = InStr(eng_lyrics, vbCrLf & vbCrLf)
    
        ' Extract the first line
        If posFirstLineEnd > 0 Then
            citation = Left(eng_lyrics, posFirstLineEnd - 1)
            If InStr(citation, vbCrLf) > 0 Or Len(citation) < 6 Then
                citation = ""
            Else
                eng_lyrics = Mid(eng_lyrics, posFirstLineEnd + Len(vbCrLf) * 2)
            End If
        End If
        ' Extract the remaining text after removing the first line
        
        
        ' Split text into slides based on double newline
        slides = Split(eng_lyrics, vbCrLf & vbCrLf)
        hin_slides = Split(hin_lyrics, vbCrLf & vbCrLf)
        
        es = UBound(slides) - LBound(slides) + 1
        hs = UBound(hin_slides) - LBound(hin_slides) + 1
        Debug.Print "eng slides " & UBound(slides) - LBound(slides) + 1
        Debug.Print "hindi slides " & UBound(hin_slides) - LBound(hin_slides) + 1
        
        If eng_lyrics <> "" And hin_lyrics <> "" And es <> hs Then
            MsgBox "English and Hindi slides for song " & songName & " do not have the same number of slides, therefor this song will not be added to the PPT. Please ensure the number of slides are equal and re-run this program to include this song."
            GoTo SkipSong
        End If
            
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
            textboxHeight = slideHeight - 24
            If eng_lyrics = "" Then
                center = 12
                textboxWidth = SlideWidth - 24
            ElseIf hin_lyrics = "" Then
                textboxWidth = SlideWidth - 24
            Else
                center = SlideWidth / 2 + 12
                textboxWidth = SlideWidth / 2 - 24
            End If
                
            
            If eng_lyrics <> "" Then
                ' Add first textbox
                Set pptTextbox1 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 12, 18, textboxWidth, textboxHeight)
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
            End If
            
    
            If hin_lyrics <> "" Then
                ' Add second textbox
                Set pptTextbox2 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, center, 18, textboxWidth, textboxHeight)
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
            End If
            
            Set footerShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 18, textboxHeight - 36, SlideWidth, slideHeight - 36)
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
            
SkipSong:
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
    
    Open "C:\Users\admin\Desktop\Song list.txt" For Input As #1
    songs = Input$(LOF(1), 1)
    Close #1
    songNames = Split(songs, vbCrLf)
    ' Define the list of .txt files (replace with your file paths)
    ' songNames = Array("Lyrics", "Lyrics2")
    
    ' Define font properties
    titleFontName = "Arial"
    titleFontSize = 60
    lyricsFontName = "Times New Roman"
    lyricsFontSize = 43
    
    ' Call the main function to create the presentation
    CreatePresentationFromTextFiles songNames, titleFontName, titleFontSize, lyricsFontName, lyricsFontSize
End Sub

