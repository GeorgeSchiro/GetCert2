param
(
    [string] $PausePrompt = "Press any key to continue...",
    [string[]] $Message = @()
)

$ScriptsIniPathFile = Join-Path -Path $env:SystemRoot -ChildPath "System32\GroupPolicy\Machine\Scripts\scripts.ini"

$Stack = Get-PSCallStack

Write-Output ""

if ( $Stack.Count -lt 2 -or -not $Stack[1].ScriptName )
{
    Write-Output $MyInvocation.MyCommand.Name
    Write-Output ""
}

Write-Output "Startup scripts in: $ScriptsIniPathFile"
Write-Output ""

if ( -not (Test-Path -LiteralPath $ScriptsIniPathFile) )
{
    Write-Output "    (there is no startup script policy file)"
}
else
{
    $InStartup = $false
    $Found = $false

    foreach ( $Line in @(Get-Content -LiteralPath $ScriptsIniPathFile) )
    {
        $Trimmed = $Line.Trim()

        if ( $Trimmed -like "[[]*[]]" )
        {
            $InStartup = $Trimmed -eq "[Startup]"
            continue
        }

        if ( $InStartup -and $Trimmed -match '^(\d+)CmdLine\s*=\s*(.*)$' )
        {
            $Found = $true

            Write-Output ("    {0}: {1}" -f $Matches[1], $Matches[2].Trim())
        }
    }

    if ( -not $Found )
    {
        Write-Output "    (no startup scripts are registered)"
    }
}

if ( $Message.Count )
{
    Write-Output ""

    foreach ( $Line in $Message )
    {
        Write-Output $Line
    }
}

Write-Output ""
Write-Output $PausePrompt

[System.Console]::ReadKey($true) | Out-Null
