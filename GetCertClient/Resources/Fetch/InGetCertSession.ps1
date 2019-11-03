param
(
    [string]$ScriptPathFile = $null
)

if ( [string]::IsNullOrEmpty($ScriptPathFile) )
{
    Write-Error "No script was passed to the GetCert session."
}
else
{
    $GetCertSession = Connect-PSSession -ComputerName localhost -Name GetCert

    Invoke-Command -Session $GetCertSession -ScriptBlock { & $args[0] } -ArgumentList $ScriptPathFile

    $discard = Disconnect-PSSession -Name GetCert
}
