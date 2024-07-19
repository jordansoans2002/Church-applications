Add-Type -AssemblyName System.Windows.Forms

$outputFile = "Song list.txt"
Write-Host "Enter list of songs"
$songList = @()

while ($true) {
	[System.Windows.Forms.SendKeys]::SendWait("YourText")
	$songName = Read-Host "Enter song name"
	if ($songName -eq ""){
		break
	}

	$lang1 = Read-Host "Enter language 1"
	$lang2 = Read-Host "Enter language 2"

	$songList += "$songName`n$lang1`n$lang2`n"
}

$songList | Set-Content $outputFile


Add-type -AssemblyName office
$application = New-Object -ComObject PowerPoint.Application
$presentation = $application.Presentations.Open("C:\Users\admin\Desktop\Church\Macro presentation.pptm")
$presentation.Application.Run("Macro presentation.pptm!RunPresentationCreation")
Start-Sleep -s 5
$presentation.Close()
