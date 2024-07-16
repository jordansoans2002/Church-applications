Sub CreatePresentationFromTextFiles(songNames As Variant, config As Variant)
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
        filePath = config("songLyricsPath") + songName
        
        ' Check if the file exists
        If Dir(filePath + "_" + config("lang1") + ".txt") <> "" Then
            ' Read the entire text from the file
            Open filePath + "_" + config("lang1") + ".txt" For Input As #1
            lyrics_lang1 = Input$(LOF(1), 1)
            Close #1
            
            cp1 = InStr(lyrics_lang1, vbCrLf & vbCrLf)
            c1 = Left(lyrics_lang1, cp1 - 1)
        Else
            MsgBox config("lang1") + " Lyrics were not found for " & songName & ". Please ensure lyrics for the song are present in the songs folder and songname entered is correct and matches the lyrics file in the folder."
            lyrics_lang1 = ""
            c1 = ""
        End If
        
         ' Check if the file exists
        If Dir(filePath + "_" + config("lang2") + ".txt") <> "" Then
             ' Read the entire text from the file
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set ts = fs.OpenTextFile(filePath + "_" + config("lang2") + ".txt", 1, False, -1) ' Open file for reading with specified encoding
            lyrics_lang2 = ts.ReadAll
            ts.Close
            Set ts = Nothing
            Set fs = Nothing
            
            cp2 = InStr(lyrics_lang2, vbCrLf & vbCrLf)
            c2 = Left(lyrics_lang2, cp2 - 1)
        Else
            If lang2 <> "" Then
                MsgBox config("lang2") + " Lyrics were not found for " & songName & ". Please ensure lyrics for the song are present in the songs folder and song name entered is correct and matches the lyrics file in the folder."
            End If
            lyrics_lang2 = ""
            c2 = ""
        End If
        
        citation = ""
        If c1 <> "" And InStr(c1, vbCrLf) = 0 Then
            citation = c1
            lyrics_lang1 = Mid(lyrics_lang1, cp1 + Len(vbCrLf) * 2)
        End If
        If c2 <> "" And InStr(c2, vbCrLf) = 0 Then
            citation = c2
            lyrics_lang2 = Mid(lyrics_lang2, cp2 + Len(vbCrLf) * 2)
        End If
        
        ' Split text into slides based on double newline
        l_text = Split(lyrics_lang1, vbCrLf & vbCrLf)
        r_text = Split(lyrics_lang2, vbCrLf & vbCrLf)
        
        If lyrics_lang1 = "" Then
            Text = r_text
        Else
            Text = l_text
        End If
        
        l1 = UBound(l_text) - LBound(l_text) + 1
        l2 = UBound(r_text) - LBound(r_text) + 1
        Debug.Print "eng slides " & UBound(l_text) - LBound(l_text) + 1
        Debug.Print "hindi slides " & UBound(r_text) - LBound(r_text) + 1
        
        If lyrics_lang1 <> "" And lyrics_lang2 <> "" And l1 <> l2 Then
            MsgBox config("lang1") + " and " + config("lang2") + "slides for song " & songName & " do not have the same number of slides, therefor this song will not be added to the PPT. Please ensure the number of slides are equal and re-run this program to include this song."
            GoTo SkipSong
        End If
            
        ' Create title slide for the song
        Set titleSlide = pptPres.slides.Add(pptPres.slides.Count + 1, 1) ' ppLayoutTitle = 1
        
        ' Set title slide text and format
        titleSlide.Shapes.Title.textFrame.TextRange.Text = songName
        With titleSlide.Shapes.Title.textFrame.TextRange.Font
            .Name = config("titleFontName")
            .Size = config("titleFontSize")
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
            .Name = config("lyricsFontName")
            .Size = 18
        End With
        ' Center the lyrics horizontally and vertically
        With footerShape.textFrame.TextRange.ParagraphFormat
            .Alignment = 1
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        
        ' Loop through each slide text and add to presentation
        For slideIndex = LBound(Text) To UBound(Text)
            ' Create a new slide with two-column text layout
            Set pptSlide = pptPres.slides.Add(pptPres.slides.Count + 1, ppLayoutBlank) ' ppLayoutTwoColumnText = 3
            SlideWidth = pptSlide.Master.Width
            slideHeight = pptSlide.Master.Height
            
            ' Calculate dimensions for textboxes
            textboxHeight = slideHeight - 36
            If lyrics_lang1 = "" Then
                center = 12
                textboxWidth = SlideWidth - 24
            ElseIf lyrics_lang2 = "" Then
                textboxWidth = SlideWidth - 24
            Else
                center = SlideWidth / 2 + 12
                textboxWidth = SlideWidth / 2 - 24
            End If
                
            
            If lyrics_lang1 <> "" Then
                ' Add first textbox
                Set pptTextbox1 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 12, 18, textboxWidth, textboxHeight)
                pptTextbox1.textFrame.TextRange.Text = l_text(slideIndex)
                ' Set lyrics slide text in the left text box
                With pptTextbox1.textFrame.TextRange.Font
                    .Name = config("lyricsFontName")
                    .Size = config("lyricsFontSize")
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
            
    
            If lyrics_lang2 <> "" Then
                ' Add second textbox
                Set pptTextbox2 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, center, 18, textboxWidth, textboxHeight)
                pptTextbox2.textFrame.TextRange.Text = r_text(slideIndex)
                 ' Set lyrics slide text in the left text box
                With pptTextbox2.textFrame.TextRange.Font
                    .Name = "Mangal"
                    .Size = config("lyricsFontSize") - 2
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
            
            Set footerShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 18, slideHeight - 36, SlideWidth, 36)
            footerShape.textFrame.TextRange.Text = citation
             ' Set lyrics slide text in the left text box
            With footerShape.textFrame.TextRange.Font
                .Name = config("lyricsFontName")
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
    Set fs = Nothing
    Set ts = Nothing
End Sub


Public Sub RunPresentationCreation()
    Dim songNames As Variant

    configPath = "C:\Users\admin\Desktop\Church\config.txt"
    Set config = CreateObject("Scripting.Dictionary")
    
    f = FreeFile
    Open configPath For Input As f
    Do While Not EOF(f)
        Line Input #f, Line
        sepPos = InStr(Line, "=")
        If sepPos > 0 Then
            key = Trim(Left(Line, sepPos - 1))
            Value = Trim(Mid(Line, sepPos + 1))
            
            config.Add key, Value
        End If
    Loop
    Close f
       
    Open config("songListPath") For Input As #1
    songs = Input$(LOF(1), 1)
    Close #1
    songNames = Split(Trim(songs), vbCrLf)
    
    ' Call the main function to create the presentation
    CreatePresentationFromTextFiles songNames, config
End Sub
