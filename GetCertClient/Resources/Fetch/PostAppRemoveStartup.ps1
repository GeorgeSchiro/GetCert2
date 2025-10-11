$RegistryName = "AutoCert\Startup"
Get-ItemProperty    -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $RegistryName
& .\PostAppShowStartups.ps1
