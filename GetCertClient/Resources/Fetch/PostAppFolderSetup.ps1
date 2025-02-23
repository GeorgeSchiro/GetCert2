param
(
    [string] $ExePath
)

if ( $ExePath -eq "" )
{
    exit
}

$ScriptPathFile = (Get-Item -Path (Join-Path -Path $ExePath -ChildPath "Startup.cmd")).FullName

if ( -not [String]::IsNullOrEmpty($ScriptPathFile) )
{
    Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "GoGetCert\Startup" -Value "`"$ScriptPathFile`""
}
