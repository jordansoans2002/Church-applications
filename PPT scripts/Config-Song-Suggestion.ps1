$global:config = @{}
$global:lyricsPath = ""
$global:songList = @()
$global:language1 = ""
$global:language2 = ""
$global:orientation = ""

function Get-Languages($songName,$n){
    $languages = @()
    $suggestions = @()

	Write-Host "`nPlease select a language. If any song is not present in chosen language it will not be added"
    Get-ChildItem -Path $global:lyricsPath -Name -File -Filter "$songName*" | ForEach-Object {
        if($_.LastIndexOf("_") -gt 0){
            $len = $_.LastIndexOf(".") - $_.LastIndexOf("_") - 1
            $nm = $_.Substring($_.LastIndexOf("_")+1,$len)

            if ($Suggestions -notcontains $nm){
                $Suggestions += $nm
                Write-Host "$($suggestions.Length). $nm"
            }
        }
    }

    $c1 = Read-Host "Language 1 selected "
    if($n -lt 2 -and $c1 -gt 0 -and $c1 -le $suggestions.Length){
        $languages += $suggestions[$c1-1]
    } else {
        Return $languages
    }
    if($n -eq 1){
        $c2 = Read-Host "Language 2 selected "
        if ($c1 -ne $c2 -and $c2 -gt 0 -and $c2 -le $suggestions.Length){
            $languages += $suggestions[$c2-1]
        }
    }
    Return $languages
}

function Show-Menu($options,$default){
    for($i=1;$i -le $options.Length;$i++){
        Write-Host "$i. $($options[$i-1])"
    }
    $opt = Read-Host "Option selected "
    if($opt -gt 0 -and $opt -le $options.Length){
        Return $opt-1
    } else{
        Return $default
    }
}

function Show-Songs-Settings(){
	Clear-Host
    if($global:language1){
		Write-Host "Language 1: $global:language1"
		if($global:language2){
			Write-Host "Language 2: $global:language2"
			if($global:orientation){
				Write-Host "Orientation: $global:orientation"
			}
		}
    }

	foreach($song in $global:songList){
		Write-Host "`nSong Name: $($song.Name)"
		if(-not $global:language1){
			Write-Host "Language 1: $($song.Languages[0])"
			if($song.Languages.Length -eq 2){
				Write-Host "Language 2: $($song.Languages[1])"

				if(-not $global:orientation){
					Write-Host "Orientation: $($song.Orientation)"
				}
			}
		}
	}
	Write-Host ""
}

function Get-Song(){
	$in = ""
	$pos = ""
	$Suggestions = @()

	while($true){
		$Suggestions = @()
		Get-ChildItem -Path $global:lyricsPath -Name -File -Filter "*$in*_*" | ForEach-Object {
			if($_.LastIndexOf("_") -gt 0){
				$nm = $_.Substring(0,$_.LastIndexOf("_"))

				if ($Suggestions -notcontains $nm){
					$Suggestions += $nm
				}
			}
		}

		#print all the songs entered uptil now with their languages
		Show-Songs-Settings

		if ($Suggestions.Length -gt 0) {
			Write-Host "Available songs:"

			for($i=1;$i -le $Suggestions.Length;$i++){
				if($pos -ne 0 -and $pos -eq "" -and $i -lt 9){$p="$i. "}
				else{$p=""}
				Write-Host $p$($Suggestions[$i-1])
			}
			Write-Host ""

			if($pos -ne 0 -and $pos -eq "" -or $pos -ge $Suggestions.Length){
				Write-Host -NoNewline "Song name: $in"
				$pos = ""
			} else{
				Write-Host -NoNewline "Song name: $($Suggestions[$pos])"
			} 
		} else {
            if($in){Write-Host "This song is not available please check the song name or add the song to the songs folder"}
            else{Write-Host "Please enter the name of your song"}
            Write-Host -NoNewline "Song name: $in"
		}

		do{
			$c = $Host.UI.RawUI.ReadKey("IncludeKeyDown,NoEcho")
			# Write-Host $c
			switch($c.VirtualKeycode){
                # letters and space
				{$_ -ge 65 -and $_ -le 90 -or $_ -eq 32}{
					if ($pos -ne ""){
						$in = $Suggestions[$pos] + $c.Character
						$pos = ""
					} else{
						$in += $c.Character
					}
					Break
				} # numbers from 1 to 9 (menu selection)
				{$_ -ge 49 -and $_ -le 57 -or $_ -ge 97 -and $_ -le 105}{
					if($pos -ne 0 -and $pos -eq "" -and $c.Character -le "$($suggestions.Length)"){
						$pos = "$($c.Character)"/1 -1
					}
					Break
				} #backspace
				{$_ -eq 8}{
					if ($pos -eq 0 -or $pos -ne ""){
						$in = $Suggestions[$pos].Substring(0,$Suggestions[$pos].Length-1)
						$pos = ""
					} elseif($in.Length -gt 0){
						$in = $in.Substring(0,$in.Length-1)
					}
					Break
				} # up arrow
				{$_ -eq 38 -and $Suggestions.Length -gt 0}{
					if($pos -ne 0 -and $pos -eq ""){
						$pos = $Suggestions.Length -1
					} elseif($pos -eq 0){
						$pos = $Suggestions.Length + 1
					} else{
						$pos -= 1
					}
					Break
				} #down arrow
				{$_ -eq 40 -and $Suggestions.Length -gt 0}{
					if($pos -ne 0 -and $pos -eq ""){
						$pos = 0
					} else{
						$pos += 1
					}
					Break
				} # enter
				{$_ -eq 13}{
					if($pos -ne "" -or $pos -eq 0){
						Return  $Suggestions[$pos]
					}elseif($pos -ne 0 -and $pos -eq "" -and $in){
						$Suggestions | ForEach-Object {
							if($_ -eq $in){
                                Return $in
                            }
						}
						$in = ""
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


# try to put relative path (./config.txt)
Get-Content "./config.txt" | Foreach-Object{
    # $configText = $_.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
    $key, $value = $_.Split("=")
    if($key -and $value){
         $global:config[$key.Trim()] = $value.Trim()
    }
}
$hymn = Read-Host "Press 'Y' to make Hymn ppt, otherwise press any other key"
if($hymn -eq "Y"){
	$global:lyricsPath = $global:config["hymnLyricsPath"] 
} else {
	$global:lyricsPath = $global:config["songLyricsPath"]
}

$langSettings = @(
    "Single language for all songs"
    "Same two languages for all songs"
    "Manually add language for each song"
)
Write-Host "`nPlease enter language settings for the presentation"
$langSetting = Show-Menu $langSettings 2
if($langSetting -ne 2){
    $languages = Get-Languages "" $langSetting
    if ($languages){
        if($languages.Length -eq 2){
            $global:language1 = $languages[0]
            $global:language2 = $languages[1]
        } else {
            $global:language1 = $languages
        }
    }
}

if($langSetting -eq 1 -or $langSetting -eq 2){
	$orientationSettings = @(
		"Stack vertical for all songs"
        "Stack side by side for all songs"
        "Manually add orientation for each song"
		)
	Write-Host "`nPlease enter orientation settings for the presentation"
    $c = Show-Menu $orientationSettings 2
    if($c -eq 0){
        $global:orientation = "vertical"
    } elseif($c -eq 1){
        $global:orientation = "horizontal"
    }
}


while($true){
	$songName = Get-Song
	if(-not $songName){Break}
	$song = @{
		Name = $songName
		Languages = @()
	}
	if($global:language1){
		$song.Languages += $global:language1
		if($global:language2){
			$song.Languages += $global:language2
			if($global:orientation){
				$song['Orientation'] = $global:orientation
			} else {
                $orientationSettings = @(
                    "Stack vertical"
                    "Stack side by side"
                )
				Write-Host "`nPlease enter orientation settings for this song"
                $c = Show-Menu $orientationSettings 1
                if($c -eq 0){
                    $song['Orientation'] = "vertical"
                } elseif($c -eq 1){
                    $song['Orientation'] = "horizontal"
                }
			}
		}
	} else {
		$languages = Get-Languages $song.Name 1
        if ($languages){
			$song.Languages += $languages

			if($song.Languages.Length -gt 1){
				if($global:orientation){
					$song['Orientation'] = $global:orientation
				} else {
					$orientationSettings = @(
						"Stack vertical"
						"Stack side by side"
					)
					Write-Host "`nPlease enter orientation settings for this song"
					$c = Show-Menu $orientationSettings 1
					if($c -eq 0){
						$song['Orientation'] = "vertical"
					} elseif($c -eq 1){
						$song['Orientation'] = "horizontal"
					}
				}
			}
        } else {
			$song = $null
		}		
	}
	if($song){
		$global:songList += $song
	}
}
Show-Songs-Settings

& "PPT Scripts\Create-PPT.ps1" -hymn $hymn -songList $global:songList