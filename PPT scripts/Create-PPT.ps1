Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

# Function to read the song list file
function Read-SongList($filePath) {
    $songList = @()
    $lines = Get-Content $filePath
    for ($i = 0; $i -lt $lines.Count; $i += 3) {
        $song = @{
            Name = $lines[$i]
            Orientation = $lines[$i + 1]
            Languages = @($lines[$i + 2])
        }
        if ($i + 2 -lt $lines.Count -and $lines[$i + 2] -match '^[A-Za-z0-9]') {
            $song.Languages += $lines[$i + 2]
            $i++
        }
        $songList += $song
    }
    return $songList
}

# Function to check if lyrics files exist
function Check-LyricsExist($songName, $languages, $lyricsFolder) {
    foreach ($lang in $languages) {
        $fileName = "${songName}_${lang}.txt"
        $filePath = Join-Path $lyricsFolder $fileName
        if (-not (Test-Path $filePath)) {
            [System.Windows.Forms.MessageBox]::Show("Error: Lyrics file not found for '$songName' in language '$lang'.", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
    }
    return $true
}

# Function to read lyrics from file and split it into slides by newlines
function Read-Lyrics($songName, $language, $lyricsFolder) {
    $fileName = "${songName}_${language}.txt"
    $filePath = Join-Path $lyricsFolder $fileName
    $lyrics = Get-Content $filePath
    $slides = @()
    $slideText = ""
    for ($i = 0; $i -lt $lyrics.Count; $i++) {
        if($lyrics[$i] -eq "" -or $i -eq $lyrics.Count-1){
            # We've encountered an empty line, add the current slide and start a new one
            $slides += $slideText.Trim()
            $slideText = ""
        } else {
            $slideText += $lyrics[$i] + "`n"
        }
    }
    return $slides
}

# Function to create a new PowerPoint presentation
function New-PowerPointPresentation {
    $application = New-Object -ComObject PowerPoint.Application
    $presentation = $application.Presentations.Add()
    return $presentation
}

# Function to add a title slide for a song
function Add-TitleSlide($presentation, $songName, $citation="") {
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 1) # 1 is the layout type for Title Slide
    $slide.Shapes.Title.TextFrame.TextRange.Text = $songName
    Add-Citation-Textbox $slide $citation
    $slide.Shapes.Title.TextFrame.TextRange.Font.Name = $global:config["titleFontName"]
    $slide.Shapes.Title.TextFrame.TextRange.Font.Size = [int]$global:config["titleFontSize"]
    $slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = [int]$global:config["titleFontColor"]

    $slide.FollowMasterBackground = $false
    $slide.Background.Fill.ForeColor.RGB = [int]$global:config["titleBackground"]
}

# Function to add a slide with lyrics
function Add-LyricsSlide($presentation, $lyrics1, $lyrics2 = $null, $orientation, $citation="") {
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12) # 12 is the layout type for Blank
    Add-Citation-Textbox $slide $citation
    $slide.FollowMasterBackground = $false
    $slide.Background.Fill.ForeColor.RGB = [int]$global:config["lyricsBackground"]
    
    $slideWidth = $slide.Master.Width
    $slideHeight = $slide.Master.Height

    if ($lyrics2) {
        # Two languages
        # $orientation = 'horizontal'
        if($orientation -eq 'vertical'){
            $textBoxWidth = $slideWidth - [int]$global:config["marginHorizontal"] * 2
            $textBoxHeight = ($slideHeight - [int]$global:config["marginBottom"] - [int]$global:config["marginTop"])/2
            $x_start = [int]$global:config["marginHorizontal"]
            $y_start = ($slideHeight + [int]$global:config["marginTop"])/2
        } else {
            $textBoxWidth = ($slideWidth - [int]$global:config["marginHorizontal"] * 2)/2
            $textBoxHeight = $slideHeight - [int]$global:config["marginBottom"] - [int]$global:config["marginTop"]
            $x_start = ($slideWidth + $global:config["marginHorizontal"])/2
            $y_start = [int]$global:config["marginTop"]
        }

        $textBox1 = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal,[int]$global:config["marginHorizontal"],[int]$global:config["marginTop"],$textBoxWidth,$textBoxHeight)
        $textBox2 = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal,$x_start,$y_start,$textBoxWidth,$textBoxHeight)
        
        $textBox1.TextFrame.TextRange.Text = $lyrics1
        $textBox2.TextFrame.TextRange.Text = $lyrics2

        $textBox1.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
        $textBox2.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter

        $textBox1.TextFrame.TextRange.Font.Name = $global:config["lyricsFontName"]
        $textBox1.TextFrame.TextRange.Font.Size = [int]$global:config["lyricsFontSize"]
        $textBox1.TextFrame.TextRange.Font.Color.RGB = [int]$global:config["lang1FontColor"]

        $textBox2.TextFrame.TextRange.Font.Name = $global:config["lyricsFontName"]
        $textBox2.TextFrame.TextRange.Font.Size = [int]$global:config["lyricsFontSize"]
        $textBox2.TextFrame.TextRange.Font.Color.RGB = [int]$global:config["lang2FontColor"]
    } else {
        # Single language
        $textBox = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal,$global:config["marginHorizontal"],$global:config["marginTop"],$slideWidth - $global:config["marginHorizontal"] * 2,$slideHeight - $global:config["marginBottom"])
        $textBox.TextFrame.TextRange.Text = $lyrics1
        $textBox1.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
        $textBox1.TextFrame.TextRange.Font.Name = $global:config["lyricsFontName"]
        $textBox1.TextFrame.TextRange.Font.Size = [int]$global:config["lyricsFontSize"]
        $textBox1.TextFrame.TextRange.Font.Color.RGB = [int]$global:config["lang1FontColor"]

    }
}

function Add-Citation-Textbox($slide, $citation){
    $textBox = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal,$global:config["marginHorizontal"],$slide.Master.Height - $global:config["marginBottom"],$slide.Master.Width,$global:config['marginBottom'])
    $textBox.TextFrame.TextRange.Text = $citation
    $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignLeft
    $textBox.TextFrame.TextRange.Font.Name = $global:config["lyricsFontName"]
    $textBox.TextFrame.TextRange.Font.Size = 18
    $textBox.TextFrame.TextRange.Font.Color.RGB = [int]$global:config["lang1FontColor"]
}


$global:config = @{}
# TODO convert to relative file path and put script and config file in same directory
Get-Content "C:\Users\admin\Desktop\Church\Church-applications\config.txt" | Foreach-Object{
   $configText = $_.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
   $key, $value = $_.Split("=")
   if($key -and $value){
	$global:config[$key.Trim()] = $value.Trim()
   }
}


# Main script
$songListFile = $global:config["songListPath"]
$lyricsFolder = $global:config["songLyricsPath"]
$outputFile = "C:\Users\admin\Desktop\LyricsPresentation.pptx"

$songList = Read-SongList $songListFile

$presentation = New-PowerPointPresentation

foreach ($song in $songList) {
    if (Check-LyricsExist $song.Name $song.Languages $lyricsFolder) {
        $lyrics1 = Read-Lyrics $song.Name $song.Languages[0] $lyricsFolder
        if ($song.Languages.Count -eq 2) {
            $lyrics2 = Read-Lyrics $song.Name $song.Languages[1] $lyricsFolder
        } else { $lyrics2 =  $null }

        $citation = ""
        if ($lyrics1[0] -ne "" -and -not $lyrics1[0].contains("`n")){
            $citation = $lyrics1[0]
            $lyrics1 = $lyrics1[1..$lyrics1.Length]
        } if ($lyrics2 -and $lyrics2[0] -ne "" -and -not $lyrics2[0].contains("`n")){
            if ($citation -eq "") {
                $citation = $lyrics2[0]
            }
            $lyrics2 = $lyrics2[1..$lyrics2.Length]
        } if ($citation -eq "") {
            $citation = $song.Name
        }

        if ($lyrics1.Count -ne $lyrics2.Count){
            [System.Windows.Forms.MessageBox]::Show($song.Languages[1] + " and " + $song.Languages[2] + "slides for song " + $song.Name + " do not have the same number of slides, therefor this song will not be added to the PPT. Please ensure the number of slides are equal and re-run this program to include this song.", "Slide count mismatch", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            continue
        }

        Add-TitleSlide $presentation $song.Name $citation
        for ($i = 0; $i -lt $lyrics1.Count; $i++) {
            Add-LyricsSlide $presentation $lyrics1[$i].Trim() $(if ($lyrics2) {$lyrics2[$i].Trim()} else {$null}) $song.Orientation $citation
        }
    }
}
#$presentation.SaveAs($outputFile)
#$presentation.Close()

# [System.Windows.Forms.MessageBox]::Show("PowerPoint presentation created successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)