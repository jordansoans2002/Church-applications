Param($hymn,$songList,$startSlideShow)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Office
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint


# Function to check if lyrics files exist
function Find-Lyrics($songName, $languages, $lyricsFolder) {
    foreach ($lang in $languages) {
        $fileName = "${songName}_${lang}.txt"
        $filePath = Join-Path $lyricsFolder $fileName
        if (-not (Test-Path $filePath)) {
            Write-Host "error:Lyrics file not found for '$songName' in language '$lang'."
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
    if (Test-Path $filePath) {
        $lyrics = Get-Content $filePath
        $slides = @()
        $slideText = ""
        for ($i = 0; $i -le $lyrics.Count; $i++) {
            # if(($lyrics[$i] -eq "" -and $i -gt 0) -or $i -eq $lyrics.Count){  # seperates slides by single line
            if(($lyrics[$i] -eq "" -and $i -gt 0 -and $lyrics[$i-1] -eq "") -or $i -eq $lyrics.Count){  # seperates slides by double line
                # We've encountered an empty line, add the current slide and start a new one
                $slides += $slideText.Trim()
                $slideText = ""
            } else {
                $slideText += [string] $lyrics[$i] + "`n"
            }

            # separates the lyrics by the number of lines in a slide
            # if(($lyrics[$i] -eq "" -and $i -gt 0) -or ((($slideText -split "`n").Count) -eq 2) -or $i -eq $lyrics.Count) {
            #     $slides += $slideText.Trim()
            #     if($lyrics[$i] -ne "" -and $i -lt $lyrics.Count) {
            #         $slideText = [string] $lyrics[$i].Trim() + "`n"
            #     }
            #     else {
            #         $slideText = ""
            #     }
            # } else {
            #     $slideText += [string] $lyrics[$i].Trim() + "`n"
            # }
        }
        return ,$slides
    } else {
        Write-Host "error:Lyrics file not found for '$songName' in language '$lang'."
        return $null
    }
}

# Function to add a title slide for a song
function Add-TitleSlide($presentation, $songName, $citation="") {
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 1) # 1 is the layout type for Title Slide
    # prefix of slide name is used as option and to identify the ppts created by this script
    $slide.Name = "songPPT$($presentation.Slides.Count)"
    $slide.Shapes.Title.Name = "lyricsLang1"
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
    # prefix of slide name is used as option and to identify the ppts created by this script
    $slide.Name = "songPPT$($presentation.Slides.Count)"
    #Add-Citation-Textbox $slide $citation
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
        $textBox1.Name = "lyricsLang1"
        $textBox2 = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal,$x_start,$y_start,$textBoxWidth,$textBoxHeight)
        $textBox1.Name = "lyricsLang2"
        
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
        $textBox = $slide.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal,[int]$global:config["marginHorizontal"],[int]$global:config["marginTop"],$slideWidth - [int]$global:config["marginHorizontal"] * 2,$slideHeight - [int]$global:config["marginBottom"])
        $textBox.Name = "lyricsLang1"
        $textBox.TextFrame.TextRange.Text = $lyrics1
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
        $textBox.TextFrame.TextRange.Font.Name = $global:config["lyricsFontName"]
        $textBox.TextFrame.TextRange.Font.Size = [int]$global:config["lyricsFontSize"]
        $textBox.TextFrame.TextRange.Font.Color.RGB = [int]$global:config["lang1FontColor"]

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
Get-Content ".\config.txt" | Foreach-Object{
#    $configText = $_.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
   $key, $value = $_.Split("=")
   if($key -and $value){
	$global:config[$key.Trim()] = $value.Trim()
   }
}


# Main script
if($hymn){
    $lyricsFolder = $global:config["hymnLyricsPath"]
} else {
    $lyricsFolder = $global:config["songLyricsPath"]
}


Write-Host "If windows is not activated first open powerpoint and then run"

# my laptop doesnt allow running the latest version, the ppt created is not editable
# open powerpoint 2010 and run this script to use  powerpoint 2010
# change the version number according to requirement 
try {
    $application = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application.14")
    Write-Host "PowerPoint version found and connected."
} catch {
    try {
        $application = New-Object -ComObject PowerPoint.Application.14
        Write-Host "PowerPoint version started."
    } catch {
        Write-Host "error: Falling back to latest installed PowerPoint."
        $application = New-Object -ComObject PowerPoint.Application
    }
}

# $application = New-Object -ComObject PowerPoint.Application
$presentation = $application.Presentations.Add()

foreach ($song in $songList) {
    # logic to create languages array when script is called from backend server
    if($song.Languages.Length -gt 2){
        $langs = $song.Languages
        $song.Languages = @()
        if($langs.IndexOf(';') -gt 2){
            $song.Languages += $langs.Substring(0,$langs.IndexOf(';'))
            $song.Languages += $langs.Substring($langs.IndexOf(';')+1)
        } else {
            $song.Languages += $langs
        }

    }

    $lyrics1 = Read-Lyrics $song.Name $song.Languages[0] $lyricsFolder
    
    if ($song.Languages.Count -eq 2) {
        $lyrics2 = Read-Lyrics $song.Name $song.Languages[1] $lyricsFolder
    } else { 
        $lyrics2 =  $null 
    }

    # $citation = ""
    # if ($lyrics1[0] -ne "" -and -not $lyrics1[0].contains("`n")){
    #     $citation = $lyrics1[0]
    #     $lyrics1 = $lyrics1[1..$lyrics1.Length]
    # } if ($lyrics2 -and $lyrics2[0] -and -not $lyrics2[0].contains("`n")){
    #     if ($citation -eq "") {
    #         $citation = $lyrics2[0]
    #     }
    #     $lyrics2 = $lyrics2[1..$lyrics2.Length]
    # } if ($citation -eq "") {
    #     $citation = $song.Name
    # }

    if ($lyrics2 -and $lyrics1.Count -ne $lyrics2.Count){
        Write-Host "warning:$($song.Languages[1]) and $($song.Languages[2]) slides for song $($song.Name) do not have the same number of slides, therefor this song will not be added to the PPT. Please ensure the number of slides are equal and re-run this program to include this song."
        [System.Windows.Forms.MessageBox]::Show($song.Languages[1] + " and " + $song.Languages[2] + "slides for song " + $song.Name + " do not have the same number of slides, therefor this song will not be added to the PPT. Please ensure the number of slides are equal and re-run this program to include this song.", "Slide count mismatch", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        continue
    }

    Add-TitleSlide $presentation $song.Name $citation
    for ($i = 0; $i -lt $lyrics1.Count; $i++) {
        Add-LyricsSlide $presentation $lyrics1[$i] $(if ($lyrics2) {$lyrics2[$i]} else {$null}) $song.Orientation $citation
    }
}

if($startSlideShow){
    $presentation.SlideShowSettings.Run()
    Write-Host "200:Slide show started"
} else {
    Write-Host "200:PPT created"
}

#$presentation.SaveAs($outputFile)
#$presentation.Close()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($application) | Out-Null
Remove-Variable application

# [System.Windows.Forms.MessageBox]::Show("PowerPoint presentation created successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)