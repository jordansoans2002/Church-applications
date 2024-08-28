Param($count,$pptOption,$slideOption)
Add-type -AssemblyName office
$application = New-Object -ComObject PowerPoint.Application

$ppts = @()
$furthestSongPpt = "" # used only with "songPPT" option
foreach($pres in $application.Presentations){
    if ($pres.SlideShowWindow){
        $pos  = $pres.SlideShowWindow.View.CurrentShowPosition
        if($pptOption -eq "all" -or $pptOption -eq "single"){
            $ppts += @{
                Pres = $pres
                Pos = $pos
            }
        } elseif($pres.Slides[$pos].Name.StartsWith($pptOption)){
            $ppts += @{
                Pres = $pres
                Pos = $pos
            }
            $slideName = $pres.Slides[$pos].Name
            $slideNumber = [int]$slideName.Substring($slideName.IndexOf($pptOption)+$pptOption.Length) * $count
            if(-not $furthestSongPpt -or $slideNumber -gt $furthestSongPpt){
                $furthestSongPpt = [int]$slideNumber
            }
        }
    }   
}

if($ppts.Length -eq 0){
    if($pptOption -eq 'all' -or $pptOption-eq 'single'){
        Write-Host "503:No active presentations"
    } else {
        Write-Host "503:No active $pptOption presentations"
    }
} elseif($pptOption -eq 'single' -and $ppts.Length -gt 1) {
    Write-Host "506:More than one presentation is active. Cannot start controller in Change Single PPT mode"
} else{
    $pptEnd = $false
    foreach($ppt in $ppts){
        if ($ppt.Pos+$count -ge 1 -and $ppt.Pos+$count -le $ppt.Pres.slides.Count){
            if($furthestSongPpt){
                $gotoSlide = $ppt.Pres.Slides("$pptOption$([Math]::Abs($furthestSongPpt)+$count)").SlideIndex
                $ppt.Pres.SlideShowWindow.View.GotoSlide($gotoSlide)
            }else{
                $ppt.Pres.SlideShowWindow.View.GotoSlide($ppt.Pos + $count)
            }
            if($slideOption -eq 'text'){
                $slide = $ppt.Pres.slides[$ppt.pos+$count]
                foreach($shape in $slide.Shapes){
                    if($shape.Name.contains("lyrics")){
                        Write-Host "{`"$($shape.Name)`":`"$($shape.TextFrame.TextRange.Text)`"}"
                    }
                }
            }
        } else {
            $pptEnd = $true
        }
    }
    if($pptEnd){
        Write-Host "204:No more slides present"
    } else {
        Write-Host "200:OK"
    }
}