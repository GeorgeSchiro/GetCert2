Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
Write-Host "Press any key to continue..."
[System.Console]::ReadKey() | Out-Null
