# Kill Outlook 
Stop-Process -Name "OUTLOOK" -Force -ErrorAction SilentlyContinue

# Remove local cache files
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Outlook\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:APPDATA\Microsoft\Outlook\*" -Recurse -Force -ErrorAction SilentlyContinue

# Remove all profiles from registry
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\*" -Recurse -Force -ErrorAction SilentlyContinue

# Enable ZeroConfigExchange for auto account setup
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" `
-Name "ZeroConfigExchange" -Value 1 -PropertyType DWord -Force

# Remove First Run flag
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup" `
-Name "First-Run" -ErrorAction SilentlyContinue

# Launch Outlook to trigger new profile creation
Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"