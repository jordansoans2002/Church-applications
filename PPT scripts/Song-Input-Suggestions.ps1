$global:songList = [ordered]@{}
$global:config = @{}


function Display-Songs(){
	Clear-Host

	ForEach ($key in $global:songList.keys){
		Write-Host "Song name: $key"
		if($global:songList[$key]){
			Write-Host "Language 1: $($global:songList[$key][0])"
			if($global:songList[$key].Length -gt 1)
			{Write-Host "Language 2: $($global:songList[$key][1])"}
		}
	}
}


function Get-Song($song){
	$in = ""
	$pos = ""
	$Suggestions = @()
	
	if($song){
		$song = $song + "_"
	}

	while($true){
		$Suggestions = @()
		if($song){ $filter_str = "$song$in*" }
		else { $filter_str = "*$in*_*"}
		Get-ChildItem -Path $global:config["songLyricsPath"] -Name -File -Filter $filter_str | ForEach-Object {
			if($_.LastIndexOf("_") -gt 0){
				if($song){
					$len = $_.LastIndexOf(".") - $_.LastIndexOf("_") - 1
					$nm = $_.Substring($_.LastIndexOf("_")+1,$len)
				} else{
					$nm = $_.Substring(0,$_.LastIndexOf("_"))
				}

				if ($Suggestions -notcontains $nm){
					$Suggestions += $nm
				}
			}
		}

		#print all the songs entered uptil now with their languages
		Display-Songs

		if ($Suggestions.Length -gt 0) {
			if($song){Write-Host "Available Languages:"}
			else {Write-Host "Available songs:"}

			for($i=1;$i -le $Suggestions.Length;$i++){
				if($pos -ne 0 -and $pos -eq "" -and $i -lt 9){$p="$i. "}
				else{$p=""}
				Write-Host $p$($Suggestions[$i-1])
			}
			Write-Host ""

			if($pos -ne 0 -and $pos -eq "" -or $pos -ge $Suggestions.Length){
				if($song){Write-Host -NoNewline "Language: $in"}
				else {Write-Host -NoNewline "Song name: $in"}

				$pos = ""
			} else{
				if($song){Write-Host -NoNewline "Language: $($Suggestions[$pos])"}
				else {Write-Host -NoNewline "Song name: $($Suggestions[$pos])"}
			} 
		} else {
			if($song){
				Write-Host "This language is not available add the lyrics to the songs folder"
				Write-Host -NoNewline "Language: $in"
			}else {
				if($in){Write-Host "This song is not available please check the song name or add the song to the songs folder"}
				else{Write-Host "Please enter the name of your song"}
				Write-Host -NoNewline "Song name: $in"
			}
		}

		do{
			$c = $Host.UI.RawUI.ReadKey("IncludeKeyDown,NoEcho")
			# Write-Host $c
			switch($c.VirtualKeycode){
				{$_ -ge 65 -and $_ -le 90 -or $_ -eq 32}{
					if ($pos -ne ""){
						$in = $Suggestions[$pos] + $c.Character
						$pos = ""
					} else{
						$in += $c.Character
					}
					Break
				}
				{$_ -ge 49 -and $_ -le 57 -or $_ -ge 97 -and $_ -le 105}{
					Write-Host ($pos -eq "") ($c.Character -le "$($suggestions.Length)")
					if($pos -ne 0 -and $pos -eq "" -and $c.Character -le "$($suggestions.Length)"){
						$pos = "$($c.Character)"/1 -1
					}
					Break
				}
				{$_ -eq 8}{
					if ($pos -eq 0 -or $pos -ne ""){
						$in = $Suggestions[$pos].Substring(0,$Suggestions[$pos].Length-1)
						$pos = ""
					} elseif($in.Length -gt 0){
						$in = $in.Substring(0,$in.Length-1)
					}
					Break
				}
				{$_ -eq 38 -and $Suggestions.Length -gt 0}{
					if($pos -ne 0 -and $pos -eq ""){
						$pos = $Suggestions.Length -1
					} elseif($pos -eq 0){
						$pos = $Suggestions.Length + 1
					} else{
						$pos -= 1
					}
					Break
				}
				{$_ -eq 40 -and $Suggestions.Length -gt 0}{
					if($pos -ne 0 -and $pos -eq ""){
						$pos = 0
					} else{
						$pos += 1
					}
					Break
				}
				{$_ -eq 13}{
					if($pos -ne "" -or $pos -eq 0){
						Return  $Suggestions[$pos]
					}elseif($pos -ne 0 -and $pos -eq "" -and $in){
						$name = ""
						$Suggestions | ForEach-Object {
							if($song){
								if($_ -eq $in.Substring($in.LastIndexOf("_")+1)){$name = $_}
							}else{
								if($_ -eq $in){$name = $_}
							}
						}
						if($name){ Return $name }
						else { $in = "" }
					} else {
						Return ""
					}
				}
				default {
					$c = $false
				}
			}
		}while(-not $c)
	}
}



Get-Content "C:\Users\admin\Desktop\Church\Church-applications\config.txt" | Foreach-Object{
   $configText = $_.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
   $key, $value = $_.Split("=")
   if($key -and $value){
		$global:config[$key.Trim()] = $value.Trim()
   }
}

Write-Host "Please select one of the below options"
Write-Host "1. Single language for all songs"
Write-Host "2. Same two languages for all songs"
Write-Host "3. Manually add language for each song"
Write-Host "Select 1, 2 or 3. Press any other key to exit"
$opt = Read-Host "Option selected: "
if($opt -eq '1' -or $opt -eq '2'){
	$suggestions = @()
	Get-ChildItem -Path $global:config["songLyricsPath"] -Name | ForEach-Object {
		if($_.LastIndexOf("_") -gt 0){
			$len = $_.LastIndexOf(".") - $_.LastIndexOf("_") - 1
			$nm = $_.Substring($_.LastIndexOf("_")+1,$len)

			if ($Suggestions -notcontains $nm){
				$Suggestions += $nm
				Write-Host "$($suggestions.Length). $nm"
			}
		}
	}

	Write-Host "Please select a language. If any song is not present in chosen language it will not be added to the ppt"
	$c1 = Read-Host "Language selected: "
	if($c1 -lt $suggestions.Length){
		$lang1 = $suggestions[$c1-1]
	}
	if($opt -eq '2'){
		$c2 = Read-Host "Language selected: "
		if($c2 -ne $c1 -and $c2 -lt $suggestions.Length){$lang2 = $suggestions[$c2-1]}
	}
	
}

while($true){
	$song = Get-Song ""
	if(-not $song){Break}
	$songList[$song] = @()
	if($lang1){
		$songList[$song] += $lang1
		$songList[$song] += Get-Song $song
	} else {
		$lang1 = Get-Song $song
		if(-not $lang1){$songList.Remove($song)}
		# $songList[$song] += $global:config["defaultLang"]
	}
}

$songListFile = ""
ForEach ($key in $global:songList.keys){
	$songName = $key
	if($global:songList[$key]){
		$lang1 = $($global:songList[$key][0])
		if($global:songList[$key].Length -gt 1)
		{$lang2 = $($global:songList[$key][1])}
	}

	$songListFile += "$songName`r`n$lang1`r`n$lang2`r`n`r`n"
}

$songListFile | Set-Content $global:config["songListPath"]


Add-type -AssemblyName office
$application = New-Object -ComObject PowerPoint.Application
$presentation = $application.Presentations.Open($global:config["macroPPTPath"])
$presentation.Application.Run("Macro presentation.potm!RunPresentationCreation")
Start-Sleep -s 5
$presentation.Close()