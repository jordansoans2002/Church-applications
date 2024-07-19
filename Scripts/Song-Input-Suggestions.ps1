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

		Display-Songs
		#print all the songs entered uptil now with their languages
		if ($Suggestions.Length -gt 0) {
			if($song){Write-Host "Available Languages:"}
			else {Write-Host "Available songs:"}

			$Suggestions | ForEach-Object {Write-Host "=> $_"}
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
				{$_ -ge 48 -and $_ -le 105 -or $_ -eq 32}{
					if ($pos -ne ""){
						$in = $Suggestions[$pos] + $c.Character
						$pos = ""
					} else{
						$in += $c.Character
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



Get-Content "C:\Users\admin\Desktop\Church\config.txt" | Foreach-Object{
   $configText = $_.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
   $key, $value = $_.Split("=")
   if($key -and $value){
		$global:config[$key.Trim()] = $value.Trim()
   }
}


while($true){
	$song = Get-Song ""
	if(-not $song){Break}
	$songList[$song] = @()
	$lang1 = Get-Song $song
	if($lang1){
		$songList[$song] += $lang1
		$songList[$song] += Get-Song $song
	} else {
		$songList[$song] += $global:config["defaultLang"]
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