Param($count)
Add-type -AssemblyName office
$application = New-Object -ComObject PowerPoint.Application
$isSlideShow = $false
ForEach($pres in $application.Presentations){
    if ($pres.SlideShowWindow){
        $isSlideShow = $true
        $pos  = $pres.SlideShowWindow.View.CurrentShowPosition
        if ($pos+$count -lt $pres.slides.Count){
            $pres.SlideShowWindow.View.GotoSlide($pos + $count)
        }
    }   
}
if(-not $isSlideShow){
    Write-Host "No active presentations"
}