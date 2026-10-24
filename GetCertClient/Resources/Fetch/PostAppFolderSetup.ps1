param
(
    [string] $ExePath
)

$ErrorActionPreference = "Stop"

if ( $ExePath -eq "" )
{
    exit
}

# These GUIDs name the Scripts client side extension. Without them in "gpt.ini"
# the entry is written but never processed, which fails silently at boot.
$ScriptsExtensionGuid = "{42B5FAAE-6536-11D2-AE5A-0000F87571E3}"
$ScriptsExtension = "[$ScriptsExtensionGuid{40B6664F-4972-11D1-A7CA-0000F87571E3}]"

$PolicyFolder = Join-Path -Path $env:SystemRoot -ChildPath "System32\GroupPolicy"
$ScriptsIniPathFile = Join-Path -Path $PolicyFolder -ChildPath "Machine\Scripts\scripts.ini"
$GptIniPathFile = Join-Path -Path $PolicyFolder -ChildPath "gpt.ini"
$ScriptPathFile = (Get-Item -Path (Join-Path -Path $ExePath -ChildPath "Startup.cmd")).FullName

[void] (New-Item -Path (Split-Path -Path $ScriptsIniPathFile) -ItemType Directory -Force)

$IniLines = @()

if ( Test-Path -Path $ScriptsIniPathFile )
{
    $IniLines = @(Get-Content -Path $ScriptsIniPathFile)
}

$SectionStart = -1
$SectionEnd = $IniLines.Count

for ( $i = 0; $i -lt $IniLines.Count; $i++ )
{
    if ( $IniLines[$i].Trim() -eq "[Startup]" )
    {
        $SectionStart = $i + 1
        $SectionEnd = $IniLines.Count

        for ( $j = $SectionStart; $j -lt $IniLines.Count; $j++ )
        {
            if ( $IniLines[$j].Trim() -like "[[]*[]]" )
            {
                $SectionEnd = $j
                break
            }
        }

        break
    }
}

if ( $SectionStart -lt 0 )
{
    $IniLines += "[Startup]"
    $SectionStart = $IniLines.Count
    $SectionEnd = $IniLines.Count
}

$Section = @()

if ( $SectionEnd -gt $SectionStart )
{
    $Section = @($IniLines[$SectionStart..($SectionEnd - 1)])
}

if ( @($Section -match ('^\d+CmdLine=' + [regex]::Escape($ScriptPathFile) + '\s*$')).Count )
{
    exit
}

$NextIndex = @($Section -match '^\d+CmdLine=').Count
$NewLines = @("$($NextIndex)CmdLine=$ScriptPathFile", "$($NextIndex)Parameters=")

$Head = @()
$Tail = @()

if ( $SectionEnd -gt 0 )
{
    $Head = @($IniLines[0..($SectionEnd - 1)])
}

if ( $SectionEnd -lt $IniLines.Count )
{
    $Tail = @($IniLines[$SectionEnd..($IniLines.Count - 1)])
}

# gpedit writes these files as UTF-16, so the same encoding is used here.
Set-Content -Path $ScriptsIniPathFile -Value ($Head + $NewLines + $Tail) -Encoding Unicode

$GptLines = @()
$Version = 0

if ( Test-Path -Path $GptIniPathFile )
{
    $GptLines = @(Get-Content -Path $GptIniPathFile)
    $VersionLines = @($GptLines -match '^\s*Version\s*=\s*\d+\s*$')

    if ( $VersionLines.Count )
    {
        $Version = [int64] ($VersionLines[0] -replace '\D','')
    }

    $GptLines = @($GptLines -notmatch '^\s*Version\s*=')
}

$ExtensionLines = @($GptLines -match '^\s*gPCMachineExtensionNames\s*=')

if ( -not $ExtensionLines.Count )
{
    if ( $GptLines -notcontains "[General]" )
    {
        $GptLines = @("[General]") + $GptLines
    }

    $GptLines += "gPCMachineExtensionNames=$ScriptsExtension"
}
elseif ( $ExtensionLines[0] -notmatch [regex]::Escape($ScriptsExtensionGuid) )
{
    $GptLines = @($GptLines -notmatch '^\s*gPCMachineExtensionNames\s*=') + ($ExtensionLines[0].Trim() + $ScriptsExtension)
}

# Adding 1 increments the computer half of the version (the low 16 bits) and leaves the user half alone.
$GptLines += "Version=$($Version + 1)"

Set-Content -Path $GptIniPathFile -Value $GptLines -Encoding Unicode

$ProfilePathFile = ""

foreach ( $ProfileFile in @(Get-ChildItem -Path $ExePath -Filter "*.exe.txt" -File) )
{
    if ( @(Get-Content -Path $ProfileFile.FullName) -match '^\s*-BackupSet\s*=\s*\[' )
    {
        $ProfilePathFile = $ProfileFile.FullName
        break
    }
}

if ( $ProfilePathFile -eq "" )
{
    exit
}

$ProfileBytes = [IO.File]::ReadAllBytes($ProfilePathFile)
$ProfileLines = @(Get-Content -Path $ProfilePathFile)

$Blocks = @()
$i = 0

while ( $i -lt $ProfileLines.Count )
{
    if ( $ProfileLines[$i] -notmatch '^\s*-(\w+)\s*=\s*\[\s*$' )
    {
        $i++
        continue
    }

    $BlockName = $Matches[1]
    $BlockEnd = $ProfileLines.Count - 1

    for ( $j = $i + 1; $j -lt $ProfileLines.Count; $j++ )
    {
        if ( $ProfileLines[$j] -match ('^\s*-' + $BlockName + '\s*=\s*\]\s*$') )
        {
            $BlockEnd = $j
            break
        }
    }

    $Blocks += @{ Name = $BlockName; Start = $i; End = $BlockEnd }
    $i = $BlockEnd + 1
}

$DesktopFolder = [Environment]::GetFolderPath("DesktopDirectory").TrimEnd('\')
$OutputFilenames = @()
$RemoveLines = @{}

foreach ( $Block in $Blocks )
{
    if ( $Block.Name -ne "BackupSet" )
    {
        continue
    }

    $IsDesktopSet = $false
    $SetOutputFilename = ""

    for ( $j = $Block.Start + 1; $j -lt $Block.End; $j++ )
    {
        if ( $ProfileLines[$j] -match '^\s*-FolderToBackup\s*=\s*(.*)$' )
        {
            $Folder = $Matches[1].Trim().Trim('"').TrimEnd('\')

            if ( $Folder -eq $DesktopFolder -or $Folder.StartsWith($DesktopFolder + '\', "OrdinalIgnoreCase") )
            {
                $IsDesktopSet = $true
            }
        }
        elseif ( $ProfileLines[$j] -match '^\s*-OutputFilename\s*=\s*(.*)$' )
        {
            $SetOutputFilename = $Matches[1].Trim().Trim('"')
        }
    }

    if ( $IsDesktopSet )
    {
        for ( $j = $Block.Start; $j -le $Block.End; $j++ )
        {
            $RemoveLines[$j] = $true
        }

        if ( $SetOutputFilename -ne "" )
        {
            $OutputFilenames += $SetOutputFilename
        }
    }
}

foreach ( $Block in $Blocks )
{
    if ( $Block.Name -ne "CleanupSet" )
    {
        continue
    }

    foreach ( $OutputFilename in $OutputFilenames )
    {
        $Purges = $false

        for ( $j = $Block.Start + 1; $j -lt $Block.End; $j++ )
        {
            if ( $ProfileLines[$j] -match '^\s*-FilesToDelete\s*=\s*(.*)$' )
            {
                $FileSpec = $Matches[1].Trim().Trim('"')

                if ( $FileSpec -ne "" -and (Split-Path -Path $FileSpec -Leaf).StartsWith($OutputFilename, "OrdinalIgnoreCase") )
                {
                    $Purges = $true
                }
            }
        }

        if ( $Purges )
        {
            for ( $j = $Block.Start; $j -le $Block.End; $j++ )
            {
                $RemoveLines[$j] = $true
            }
        }
    }
}

if ( -not $RemoveLines.Count )
{
    exit
}

$KeptLines = @()

for ( $i = 0; $i -lt $ProfileLines.Count; $i++ )
{
    if ( -not $RemoveLines.ContainsKey($i) )
    {
        $KeptLines += $ProfileLines[$i]
    }
}

$ProfileLines = $KeptLines

$ProfileEncoding = New-Object System.Text.UTF8Encoding($false)

if ( $ProfileBytes.Length -ge 2 -and $ProfileBytes[0] -eq 0xFF -and $ProfileBytes[1] -eq 0xFE )
{
    $ProfileEncoding = New-Object System.Text.UnicodeEncoding($false, $true)
}
elseif ( $ProfileBytes.Length -ge 3 -and $ProfileBytes[0] -eq 0xEF -and $ProfileBytes[1] -eq 0xBB -and $ProfileBytes[2] -eq 0xBF )
{
    $ProfileEncoding = New-Object System.Text.UTF8Encoding($true)
}

[IO.File]::WriteAllText($ProfilePathFile, (($ProfileLines -join "`r`n") + "`r`n"), $ProfileEncoding)
