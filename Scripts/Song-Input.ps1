$outputFile = "Song list.txt"

$songList = @()

while ($true) {
	$songName = Read-Host "Enter song name"

	if ($songName -eq ""){
		break
	}

	$songList += $songName
}

Add-type -AssemblyName office

$application = New-Object -ComObject PowerPoint.Application
$presentation = $application.Presentations.Open("C:\Users\admin\Desktop\Church\Macro presentation.pptm")
$presentation.Application.Run("Macro presentation.pptm!RunPresentationCreation")
Start-Sleep -s 2
$presentation.Close()

$songList | Out-File -FilePath $outputFile -Encoding UTF8