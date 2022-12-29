$usern = [System.Environment]::UserName

Stop-Process -Name "Outlook" -Force

cd C:\Users\$usern\Appdata\Local\Microsoft\Outlook

Get-ChildItem -Filter "OLD - $usern*" | Remove-Item

Start-Sleep -m 1000

Get-ChildItem -Filter "$usern*" | Rename-Item -NewName {$_.Name -replace "^", "OLD - "}

Start-Process Outlook.exe