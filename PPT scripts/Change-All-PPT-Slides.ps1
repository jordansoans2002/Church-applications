$n = $args[0]
Add-type -AssemblyName office
$application = New-Object -ComObject PowerPoint.Application
ForEach($pres in $application.Presentations){
    if ($pres.SlideShowWindow){
        $pos  = $pres.SlideShowWindow.View.CurrentShowPosition
        if ($pos -lt $pres.slides.Count){
            $pres.SlideShowWindow.View.GotoSlide($pos + $n)
        }
    }   
}