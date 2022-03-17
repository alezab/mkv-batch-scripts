# Set location to script directory
Set-Location $PSScriptRoot
# Load Get-CRC32 function
. .\crc32.ps1
# Variables declaration
$Conf = ($PSScriptRoot + "\conf.json")
$MKVmerge = (Get-Content $Conf | ConvertFrom-Json).mkvtoolnix_path + "mkvmerge.exe"

function ListShows {
    Write-Output "`nList of shows:"
    $x = 1
    foreach ($Folder in (Get-ChildItem ".\shows")) {
        Write-Output ("" + $x + ". " + $Folder.Name)
        $x++
    }
    
}

function ListMovies {
    Write-Output "`nList of movies:"
    $x = 1
    foreach ($Folder in (Get-ChildItem ".\movies")) {
        Write-Output ("" + $x + ". " + $Folder.Name)
        $x++
    }       
}

function MergeSubs {
    param (
        [string[]]$Path,
        [string[]]$Episode,
        [string[]]$Output
    )
    $SubFlags = ""
        
    foreach ($Sub in (Get-ChildItem "$($AnimePath)\$($Episode)\*" -Include "*.ass")) { $SubFlags += " --language 0:$($AnimeProps.sub_lang) `"$($Sub.FullName)`"" }
        
    
    #Write-Host "$MKVmerge --output `"$Output`" `"$Path`"$SubFlags"

    #Invoke-Expression "&`"$MKVmerge`" --output `"$Output`" `"$Path`"$FontFlags"
    return $SubFlags
}

function MergeFonts {
    param (
        [string[]]$Path,
        [string[]]$Episode,
        [string[]]$Output
    )
    $FontFlags = ""
    
    if ([string]::IsNullOrEmpty((& $MKVmerge -J $Path | ConvertFrom-Json).attachments.file_name)) { 
        foreach ($Font in (Get-ChildItem "$($AnimePath)\$($Episode)\fonts\*" -Include "*.ttf")) { $FontFlags += " --attachment-mime-type application/x-truetype-font --attach-file `"$($Font.FullName)`"" }
        foreach ($Font in (Get-ChildItem "$($AnimePath)\$($Episode)\fonts\*" -Include "*.otf")) { $FontFlags += " --attachment-mime-type application/vnd.ms-opentype --attach-file `"$($Font.FullName)`"" }
    }
    else {
        foreach ($Font in (Get-ChildItem "$($AnimePath)\$($Episode)\fonts\*" -Include "*.ttf")) {
            if ( (& $MKVmerge -J $Path | ConvertFrom-Json).attachments.file_name.Contains($Font.Name) ) {
                Write-Output ("Attachment " + $Font.Name + " already exists in file. Skipping.") 
            }
            else { 
                $FontFlags += " --attachment-mime-type application/x-truetype-font --attach-file `"$($Font.FullName)`"" 
            }      
        }
        foreach ($Font in (Get-ChildItem "$($AnimePath)\$($Episode)\fonts\*" -Include "*.otf")) {
            if ( (& $MKVmerge -J $Path | ConvertFrom-Json).attachments.file_name.Contains($Font.Name) ) {
                Write-Output ("Attachment " + $Font.Name + " already exists in file. Skipping.") 
            }
            else { 
                $FontFlags += " --attachment-mime-type application/vnd.ms-opentype --attach-file `"$($Font.FullName)`"" 
            }      
        }
    }
    return $FontFlags
}

function MergeChapters {
    param (
        [string[]]$Path,
        [string[]]$Episode,
        [string[]]$Output
    )
    $ChapFlags = ""
        
    foreach ($Chap in (Get-ChildItem "$($AnimePath)\$($Episode)\*" -Include "*.xml")) { $ChapFlags = " --chapters `"$($Chap.FullName)`"" }

    #Write-Host "$MKVmerge --output `"$Output`" `"$Path`"$SubFlags"

    #Invoke-Expression "&`"$MKVmerge`" --output `"$Output`" `"$Path`"$FontFlags"
    return $ChapFlags
}

function MergeByHash {
    param (
        [string[]]$AnimePath,
        [array[]]$AnimeProps
    )
    foreach ($File in (Get-ChildItem "$($AnimePath)\premux\*" -Include ("*.mkv", "*.mp4"))) {
        $n = 0
        do {
            if ((Get-CRC32 -Path $File.FullName) -eq $AnimeProps.files[$n].hash) {
                Write-Output ("`nProcessing Episode " + $AnimeProps.files[$n].episode + ": File """ + $File.Name + """ with hash value of " + $AnimeProps.files[$n].hash + ".")
                Write-Output "`nDetected attachments: "
                Get-ChildItem "$AnimePath\$($AnimeProps.files[$n].episode)\*" -Include ("*.ass", "*.srt", "*.xml", "*.ttf", "*.otf") -Name -Recurse
                $Path = $File.FullName
                $Output = "$($AnimePath)\mux\$($File.Name)"       
                $SubFlags = MergeSubs -Path $File.FullName -Episode $AnimeProps.files[$n].episode -Output $Output
                $FontFlags = MergeFonts -Path $File.FullName -Episode $AnimeProps.files[$n].episode -Output $Output
                $ChapFlags = MergeChapters -Path $File.FullName -Episode $AnimeProps.files[$n].episode -Output $Output
                $Flags = $SubFlags + $FontFlags + $ChapFlags
                Write-Host "`nLaunching mkvmerge.exe with parameters:"
                Write-Host "$MKVmerge --output `"$Output`" `"$Path`"$Flags"
                #Invoke-Expression "&`"$MKVmerge`" --output `"$Output`" `"$Path`"$Flags"      
                break
            }
            $n++
        } while ($n -lt $AnimeProps.files.Count)
    }
}

function MergeByName {
    param (
        [string[]]$AnimePath,
        [array[]]$AnimeProps
    )
    $Premux = (Get-ChildItem "$($AnimePath)\premux\*" -Include ("*.mkv", "*.mp4"))
    $arr = @()
    1..$Premux.Count | ForEach-Object { $arr += $_.ToString("00") } 
    foreach ($File in $Premux) {
        $n = 0
        do {
            $Ep = $arr[$n]
            if ( (" " + $File.Name + " ") -match "[^0-9]$Ep[^0-9]") { 
                Write-Output ("`nProcessing Episode " + $Ep + ": File """ + $File.Name + """.")
                Write-Output "`nDetected attachments: "
                Get-ChildItem "$AnimePath\$Ep\*" -Include ("*.ass", "*.srt", "*.xml", "*.ttf", "*.otf") -Name -Recurse
                $Path = $File.FullName
                $Output = "$($AnimePath)\mux\$($File.Name)"
                $SubFlags = MergeSubs -Path $File.FullName -Episode $Ep -Output $Output
                $FontFlags = MergeFonts -Path $File.FullName -Episode $Ep -Output $Output
                $ChapFlags = MergeChapters -Path $File.FullName -Episode $Ep -Output $Output
                $Flags = $SubFlags + $FontFlags + $ChapFlags
                Write-Host "`nLaunching mkvmerge.exe with parameters:"
                Write-Host "$MKVmerge --output `"$Output`" `"$Path`"$Flags"
                #Invoke-Expression "&`"$MKVmerge`" --output `"$Output`" `"$Path`"$Flags"
                break
            }
            $n++
        } while ($n -lt $Premux.Count)
    }
}

Write-Output "`nChecking for mkvmerge.exe..."
if (Test-Path -Path $MKVmerge -PathType Leaf) { "Found $($MKVmerge)" } else { "mkvmerge.exe not found. Check path in $($Conf)."; exit }

function SushiSubs {
    param (
        [string[]]$SourceDir,
        [string[]]$DestinationDir,
        [string[]]$ScriptDir,
        [string[]]$SourceAudio,
        [string[]]$DestinationAudio
    )
    
    $Source = (Get-ChildItem "$($SourceDir)\*" -Include ("*.mkv", "*.mp4"))
    $SourceArray = @()
    1..$Source.Count | ForEach-Object { $SourceArray += $_.ToString("00") } 
    
    $Destination = (Get-ChildItem "$($DestinationDir)\*" -Include ("*.mkv", "*.mp4"))
    $DestinationArray = @()
    1..$Destination.Count | ForEach-Object { $DestinationArray += $_.ToString("00") } 

    $Script = (Get-ChildItem "$($ScriptDir)\*" -Include ("*.ass"))
    $ScriptArray = @()
    1..$Script.Count | ForEach-Object { $ScriptArray += $_.ToString("00") } 
    
    $n = 0
    do {
        $EpSource = $SourceArray[$n]
        $EpDestination = $DestinationArray[$n]
        $EpScript = $ScriptArray[$n]

        foreach ($File in $Source) {
            if ( (" " + $File.Name + " ") -match "[^0-9]$EpSource[^0-9]") { 
                Write-Output ("`nDetected Source Episode " + $EpSource + ": """ + $File.Name + """")
                $SourcePath = $File.FullName 
                break
            }
        }
        foreach ($File in $Destination) {
            if ( (" " + $File.Name + " ") -match "[^0-9]$EpDestination[^0-9]") { 
                Write-Output ("Detected Destination Episode " + $EpDestination + ": """ + $File.Name + """")
                $DestinationPath = $File.FullName 
                break
            }
        }
        foreach ($File in $Source) {
            if ( (" " + $File.Name + " ") -match "[^0-9]$EpScript[^0-9]") { 
                Write-Output ("Detected subtitles for Episode " + $EpScript + ": """ + $File.Name + """")
                $ScriptPath = $File.FullName 
                break
            }
        }

        #Write-Host "`nLaunching Sushi with parameters:"
        #Write-Host "sushi --src `"$SourcePath`" --dst `"$DestinationPath`" --script `"$ScriptPath`" --max-window 360 --sample-rate 48000 --src-audio `"$SourceAudio`" -dst-audio `"$DestinationAudio`""
        Invoke-Expression "& sushi --src `"$SourcePath`" --dst `"$DestinationPath`" --script `"$ScriptPath`" --max-window 360 --sample-rate 48000 --src-audio `"$SourceAudio`" --dst-audio `"$DestinationAudio`"" 
        $n++
    } while ($n -lt $Source.Count)
}

function ListTracks {
    param (
        [array[]]$File
    )
    $t = 0
    $n = 1
    do {
        if ($File.tracks.type[$t] -eq "audio") { Write-Output ("$($n). Track ID: $($File.tracks.id[$t]); Type: $($File.tracks.type[$t]); Language: $($File.tracks.properties[$t].language)"); $n++ }
        $t++
    } while ($t -lt $File.tracks.id.Count)
}
 
function StartSushi {
    $SourceDir = "$($PSScriptRoot)\sushi\source"
    $DestinationDir = "$($PSScriptRoot)\sushi\destination"
    $ScriptDir = "$($PSScriptRoot)\sushi\script"
    
    do { 
        $Confirmation = Read-Host "Use default directories? $($PSScriptRoot)\sushi\source; $($PSScriptRoot)\sushi\destination; $($PSScriptRoot)\sushi\script ) [y/n]"
        if ($Confirmation -eq "y") { break }
        
        $SourceDir = Read-Host "Choose source directory"
        $DestinationDir = Read-Host "Choose destination directory"
        $ScriptDir = Read-Host "Choose subtitles directory"
    } 
    while ($Confirmation -ne "n")
    
    $Source = (Get-ChildItem "$($SourceDir)\*" -Include ("*.mkv", "*.mp4"))
    $Destination = (Get-ChildItem "$($DestinationDir)\*" -Include ("*.mkv", "*.mp4"))
    
    $SushiSource = ( & $MKVmerge -J $Source[0].FullName | ConvertFrom-Json)
    $SushiDestination = ( & $MKVmerge -J $Destination[0].FullName | ConvertFrom-Json)
    
    do {
        $End = 0
        ListTracks -File $SushiSource
        [int]$SID = Read-Host -Prompt "Select number"
        if ($SID -gt 0 -and $SID -le $SushiSource.tracks.id.Count) { $End = 1 } else { $End = 0 }
    } while ($End -eq 0)
    
    do {
        $End = 0
        ListTracks -File $SushiDestination
        [int]$DID = Read-Host -Prompt "Select number"
        if ($DID -gt 0 -and $DID -le $SushiDestination.tracks.id.Count) { $End = 1 } else { $End = 0 }
    } while ($End -eq 0)
    
    
    SushiSubs -SourceDir $SourceDir -DestinationDir $DestinationDir -ScriptDir $ScriptDir -SourceAudio $SID -DestinationAudio $DID  
    exit
}

do {
    $End = 0
    $Cat = 0
    Write-Output "`nChoose action:"
    Write-Output "1. Merge"
    Write-Output "2. Sushi"
    [int]$Cat = Read-Host -Prompt "Select number"
    if ($Cat -gt 0 -and $Cat -lt 3) { $End = 1 } else { $End = 0 }
} while ($End -eq 0)

if ($Cat -eq 2) { StartSushi }
if ($Cat -lt 1 -or $Cat -gt 2) { throw "Error: No category found." }

$Shows = @(Get-ChildItem -Path ".\shows")
$Movies = @(Get-ChildItem -Path ".\movies")

do {
    $End = 0
    $Cat = 0
    Write-Output "`nChoose category:"
    Write-Output "1. Shows"
    Write-Output "2. Movies"
    [int]$Cat = Read-Host -Prompt "Select number"
    if ($Cat -gt 0 -and $Cat -lt 3) { $End = 1 } else { $End = 0 }
} while ($End -eq 0)

if ($Cat -lt 1 -or $Cat -gt 2) { throw "Error: No category found." }

do {
    $End = 0
    switch ($Cat) {
        1 { ListShows; [int]$Sel = Read-Host -Prompt "Select show"; if ($Sel -gt 0 -and $Sel -le $Shows.Count) { $End = 1 } else { $End = 0 }; $End; Write-Output ""; break }
        2 { ListMovies; $Sel = Read-Host -Prompt "Select movie"; Write-Output ""; break }
        Default { throw "Error: No category found." }
    }
} while ($End -eq 0)

Write-Output "Drop pre-mux video files to `"premux`" folder in root of anime selected folder."

do {
    Write-Output "`nSelect file sorting method:"
    Write-Output "1. By hash. Slower but more reliable. Only for publicly known file hashes stored in `"props.json`"."
    Write-Output "2. By name. Rename files to contain number in file name (e.g. 01, 02,..., 10, 11...n)."
    [int]$Met = Read-Host -Prompt "Select method"
    if ($Met -gt 0 -and $Met -lt 3) { $End = 1 } else { $End = 0 }
} while ($End -eq 0)

if ($Cat -eq 1 -and $Met -eq 1) { MergeByHash -AnimePath $Shows[$Sel - 1].FullName -AnimeProps (Get-Content "$($Shows[$Sel-1].FullName)\props.json" | ConvertFrom-Json) }
if ($Cat -eq 2 -and $Met -eq 1) { MergeByHash -AnimePath $Movies[$Sel - 1].FullName -AnimeProps (Get-Content "$($Movies[$Sel-1].FullName)\props.json" | ConvertFrom-Json) }
if ($Cat -eq 1 -and $Met -eq 2) { MergeByName -AnimePath $Shows[$Sel - 1].FullName -AnimeProps (Get-Content "$($Shows[$Sel-1].FullName)\props.json" | ConvertFrom-Json) }
if ($Cat -eq 2 -and $Met -eq 2) { MergeByName -AnimePath $Movies[$Sel - 1].FullName -AnimeProps (Get-Content "$($Movies[$Sel-1].FullName)\props.json" | ConvertFrom-Json) }
if ($Met -lt 1 -or $Met -gt 2) { throw "Error: No method found." }