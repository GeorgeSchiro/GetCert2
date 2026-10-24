$ErrorActionPreference = "Stop"

$Stack = Get-PSCallStack

if ( $Stack.Count -lt 2 -or -not $Stack[1].ScriptName )
{
    Write-Output ""
    Write-Output $MyInvocation.MyCommand.Name
}

$PolicyFolder = Join-Path -Path $env:SystemRoot -ChildPath "System32\GroupPolicy"
$ScriptsIniPathFile = Join-Path -Path $PolicyFolder -ChildPath "Machine\Scripts\scripts.ini"
$GptIniPathFile = Join-Path -Path $PolicyFolder -ChildPath "gpt.ini"
$ShowStartupsPathFile = Join-Path -Path $PSScriptRoot -ChildPath "PostAppShowStartups.ps1"
$ScriptPathFile = Join-Path -Path $PSScriptRoot -ChildPath "Startup.cmd"
$ConsolePathFile = Join-Path -Path $PSScriptRoot -ChildPath "Startup.msc"

$IniLines = @()

if ( Test-Path -LiteralPath $ScriptsIniPathFile )
{
    $IniLines = @(Get-Content -LiteralPath $ScriptsIniPathFile)
}

$StartupHeader = -1

for ( $i = 0; $i -lt $IniLines.Count; $i++ )
{
    if ( $IniLines[$i].Trim() -eq "[Startup]" )
    {
        $StartupHeader = $i
        break
    }
}

$SectionEnd = $IniLines.Count

for ( $i = $StartupHeader + 1; $i -lt $IniLines.Count -and $StartupHeader -ge 0; $i++ )
{
    if ( $IniLines[$i].Trim() -like "[[]*[]]" )
    {
        $SectionEnd = $i
        break
    }
}

$Entries = @{}

for ( $i = $StartupHeader + 1; $i -lt $SectionEnd -and $StartupHeader -ge 0; $i++ )
{
    $Line = $IniLines[$i].Trim()

    if ( $Line -match '^(\d+)CmdLine\s*=\s*(.*)$' )
    {
        $Index = [int] $Matches[1]

        if ( -not $Entries.ContainsKey($Index) )
        {
            $Entries[$Index] = @{ CmdLine = ""; Parameters = "" }
        }

        $Entries[$Index].CmdLine = $Matches[2].Trim()
    }
    elseif ( $Line -match '^(\d+)Parameters\s*=\s*(.*)$' )
    {
        $Index = [int] $Matches[1]

        if ( -not $Entries.ContainsKey($Index) )
        {
            $Entries[$Index] = @{ CmdLine = ""; Parameters = "" }
        }

        $Entries[$Index].Parameters = $Matches[2].Trim()
    }
}

$Kept = @()
$Removed = @()

foreach ( $Index in @($Entries.Keys | Sort-Object) )
{
    if ( $Entries[$Index].CmdLine -eq $ScriptPathFile )
    {
        $Removed += $Entries[$Index].CmdLine
    }
    else
    {
        $Kept += $Entries[$Index]
    }
}

if ( -not $Removed.Count )
{
    & $ShowStartupsPathFile -Message "There is nothing to remove."

    exit
}

$CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal(
        [Security.Principal.WindowsIdentity]::GetCurrent())

if ( -not $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) )
{
    & $ShowStartupsPathFile -Message @(
              "This must be run as an administrator to change the startup script policy."
            , ""
            , "Run 'Startup.msc' instead and remove 'Startup.cmd' yourself:"
            , ""
            , "    $ConsolePathFile"
            )

    exit
}

& $ShowStartupsPathFile -PausePrompt "Press any key to remove '$ScriptPathFile'.`n(close this console to cancel)"

$RebuiltSection = @()
$NextIndex = 0

foreach ( $Entry in $Kept )
{
    $RebuiltSection += "$($NextIndex)CmdLine=$($Entry.CmdLine)"
    $RebuiltSection += "$($NextIndex)Parameters=$($Entry.Parameters)"
    $NextIndex++
}

$Head = @($IniLines[0..$StartupHeader])
$Tail = @()

if ( $SectionEnd -lt $IniLines.Count )
{
    $Tail = @($IniLines[$SectionEnd..($IniLines.Count - 1)])
}

Set-Content -LiteralPath $ScriptsIniPathFile -Value ($Head + $RebuiltSection + $Tail) -Encoding Unicode

$GptLines = @()
$Version = 0

if ( Test-Path -LiteralPath $GptIniPathFile )
{
    foreach ( $Line in @(Get-Content -LiteralPath $GptIniPathFile) )
    {
        if ( $Line -match '^\s*Version\s*=\s*(\d+)\s*$' )
        {
            $Version = [int64] $Matches[1]
            continue
        }

        $GptLines += $Line
    }
}

if ( $GptLines -notcontains "[General]" )
{
    $GptLines = @("[General]") + $GptLines
}

# Adding 1 increments the computer half of the version (the low 16 bits) and leaves the user half alone.
$GptLines += "Version=$($Version + 1)"

Set-Content -LiteralPath $GptIniPathFile -Value $GptLines -Encoding Unicode

foreach ( $CmdLine in $Removed )
{
    Write-Output ""
    Write-Output "Removed startup script entry: $CmdLine"
}

& $ShowStartupsPathFile -Message "This takes effect at next boot."
