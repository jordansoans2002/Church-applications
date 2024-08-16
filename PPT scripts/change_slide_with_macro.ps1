Add-type -AssemblyName office
$application = New-Object -ComObject PowerPoint.Application
$presentation = $application.Presentations.Open("C:\Users\admin\Desktop\Church\Church-applications\PPT macros\Macro presentation.potm")
$presentation.Application.Run("Macro presentation.potm!NextSlideInAllPresentations")
Start-Sleep -s 5
$presentation.Close()