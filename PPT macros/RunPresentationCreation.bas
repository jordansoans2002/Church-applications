Sub CreatePresentationFromTextFiles(songs As Variant, config As Variant)
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
    
    'Set pptPres = pptApp.Presentations.Open("C:\Users\admin\Desktop\Church\Macro presentation.potm", msoFalse, msoTrue, msoFalse)
    
    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    'pptPres.ApplyTheme ("C:\Users\admin\Desktop\Church\Church-applications\PPT macros\Theme1.thmx")
    
    ' Loop through each songName passed as input
    For i = LBound(songs) To UBound(songs)
        song = Split(songs(i), vbCrLf)
        If UBound(song) - LBound(song) < 1 Then
            GoTo SkipSong
        End If
        songName = song(0)
        lang1 = song(1)
        If UBound(song) - LBound(song) > 1 Then
            lang2 = song(2)
        End If
        
        If Len(songName) < 2 Then
         GoTo SkipSong
        End If
        filePath = config("songLyricsPath") + songName
        Debug.Print filePath
        
        ' Check if the file exists
        If Dir(filePath + "_" + lang1 + ".txt") <> "" Then
             ' Read the entire text from the file
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set ts = fs.OpenTextFile(filePath + "_" + lang1 + ".txt", 1, False, -1) ' Open file for reading with specified encoding
            lyrics_lang1 = ts.ReadAll
            ts.Close
            Set ts = Nothing
            Set fs = Nothing
            
            cp1 = InStr(lyrics_lang1, vbCrLf & vbCrLf)
            c1 = Left(lyrics_lang1, cp1 - 1)
        Else
            MsgBox lang1 + " lyrics were not found for " & songName & ". Please ensure lyrics for the song are present in the songs folder and songname entered is correct and matches the lyrics file in the folder."
            lyrics_lang1 = ""
            c1 = ""
        End If
        
         ' Check if the file exists
        If Dir(filePath + "_" + lang2 + ".txt") <> "" Then
             ' Read the entire text from the file
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set ts = fs.OpenTextFile(filePath + "_" + lang2 + ".txt", 1, False, -1) ' Open file for reading with specified encoding
            lyrics_lang2 = ts.ReadAll
            ts.Close
            Set ts = Nothing
            Set fs = Nothing
            
            cp2 = InStr(lyrics_lang2, vbCrLf & vbCrLf)
            c2 = Left(lyrics_lang2, cp2 - 1)
        Else
            If lang2 <> "" Then
                MsgBox lang2 + " lyrics were not found for " & songName & ". Please ensure lyrics for the song are present in the songs folder and song name entered is correct and matches the lyrics file in the folder."
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
            MsgBox lang1 + " and " + lang2 + "slides for song " & songName & " do not have the same number of slides, therefor this song will not be added to the PPT. Please ensure the number of slides are equal and re-run this program to include this song."
            GoTo SkipSong
        End If
            
        ' Create title slide for the song
        Set titleSlide = pptPres.slides.Add(pptPres.slides.Count + 1, 1) ' ppLayoutTitle = 1
        titleSlide.FollowMasterBackground = msoFalse
        titleSlide.Background.Fill.ForeColor.RGB = RGB(Split(config("titleBackground"), ",")(0), Split(config("titleBackground"), ",")(1), Split(config("titleBackground"), ",")(2))
        
        ' Set title slide text and format
        titleSlide.Shapes.Title.textFrame.TextRange.Text = songName
        With titleSlide.Shapes.Title.textFrame.TextRange.Font
            .Name = config("titleFontName")
            .Size = config("titleFontSize")
            .Color.RGB = RGB(Split(config("titleFontColor"), ",")(0), Split(config("titleFontColor"), ",")(1), Split(config("titleFontColor"), ",")(2))
        End With
        
        ' Center title text horizontally and vertically
        With titleSlide.Shapes.Title.textFrame.TextRange.ParagraphFormat
            .Alignment = 2 ' ppAlignCenter = 2
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        
        Debug.Print config("lang1Color")
        Set footerShape = titleSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, config("marginHorizontal"), titleSlide.Master.Height - config("marginBottom"), titleSlide.Master.Width, config("marginBottom"))
        footerShape.textFrame.TextRange.Text = citation
        ' Set lyrics slide text in the left text box
        With footerShape.textFrame.TextRange.Font
            .Name = config("lyricsFontName")
            .Size = 18
            .Color.RGB = RGB(Split(config("lang1FontColor"), ",")(0), Split(config("lang1FontColor"), ",")(1), Split(config("lang1FontColor"), ",")(2))
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
            pptSlide.FollowMasterBackground = msoFalse
            pptSlide.Background.Fill.ForeColor.RGB = RGB(Split(config("lyricsBackground"), ",")(0), Split(config("lyricsBackground"), ",")(1), Split(config("lyricsBackground"), ",")(2))
'            With pptSlide.Background.Fill
'                .ForeColor.RGB = RGB(0, 0, 0)
'                .BackColor.RGB = RGB(0, 0, 0)
'            End With
                
            SlideWidth = pptSlide.Master.Width
            SlideHeight = pptSlide.Master.Height
            
            ' Calculate dimensions for textboxes
            textboxHeight = SlideHeight - config("marginBottom")
            If lyrics_lang1 = "" Then
                center = config("marginHorizontal")
                textboxWidth = SlideWidth - config("marginHorizontal") * 2
            ElseIf lyrics_lang2 = "" Then
                textboxWidth = SlideWidth - config("marginHorizontal") * 2
            Else
                center = SlideWidth / 2 + config("marginHorizontal")
                textboxWidth = SlideWidth / 2 - config("marginHorizontal") * 2
            End If
                
            
            If lyrics_lang1 <> "" Then
                ' Add first textbox
                Set pptTextbox1 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, config("marginHorizontal"), config("marginTop"), textboxWidth, textboxHeight)
                pptTextbox1.textFrame.TextRange.Text = l_text(slideIndex)
                ' Set lyrics slide text in the left text box
                
                With pptTextbox1.textFrame.TextRange.Font
                    .Name = config("lyricsFontName")
                    .Size = config("lyricsFontSize")
                    .Color.RGB = RGB(Split(config("lang1FontColor"), ",")(0), Split(config("lang1FontColor"), ",")(1), Split(config("lang1FontColor"), ",")(2))
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
                Set pptTextbox2 = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, center, config("marginTop"), textboxWidth, textboxHeight)
                pptTextbox2.textFrame.TextRange.Text = r_text(slideIndex)
                 ' Set lyrics slide text in the left text box
                With pptTextbox2.textFrame.TextRange.Font
                    .Name = "Mangal"
                    .Size = config("lyricsFontSize") - 2
                    .Color.RGB = RGB(Split(config("lang2FontColor"), ",")(0), Split(config("lang2FontColor"), ",")(1), Split(config("lang2FontColor"), ",")(2))
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
            
            Set footerShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, config("marginHorizontal"), SlideHeight - config("marginBottom"), SlideWidth, config("marginBottom"))
            footerShape.textFrame.TextRange.Text = citation
             ' Set lyrics slide text in the left text box
            With footerShape.textFrame.TextRange.Font
                .Name = config("lyricsFontName")
                .Size = 18
                .Color.RGB = RGB(Split(config("lang1FontColor"), ",")(0), Split(config("lang1FontColor"), ",")(1), Split(config("lang1FontColor"), ",")(2))
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

    configPath = "C:\Users\admin\Desktop\Church\Church-applications\config.txt"
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
    songList = Input$(LOF(1), 1)
    Close #1
    
    songs = Split(Trim(songList), vbCrLf & vbCrLf)
    
    ' Call the main function to create the presentation
    CreatePresentationFromTextFiles songs, config
End Sub
