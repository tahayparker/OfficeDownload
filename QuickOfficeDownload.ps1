# Download Office - Word, Excel, Powerpoint

# Request for Admin priveleges if not already admin to run the entire script
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Get Current Path
$CurrentPath = Get-Location

# Make a temp dir
mkdir C:\temp

# Change to the temp dir
cd C:\temp

# Download the setup file
Invoke-WebRequest https://raw.githubusercontent.com/tahayparker/OfficeDownload/refs/heads/main/setup.exe -OutFile setup.exe

# Warn user to not open the apps
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show('Office will now be installed. Please do not open any apps until another message box appears', 'Office Download', 'OK', 'Warning')

# Run the setup file
Start-Process -FilePath C:\temp\setup.exe -ArgumentList "/configure https://raw.githubusercontent.com/tahayparker/OfficeDownload/refs/heads/main/TPPublic.xml" -Wait

# Change back to the original path
cd $CurrentPath

# Remove the temp dir
Remove-Item C:\temp -Recurse -Force

# Run get.activated.win script to activate Office
& ([ScriptBlock]::Create((irm https://get.activated.win))) /Ohook /S

# Inform user via msgbox alert
[System.Windows.MessageBox]::Show('Office (Word, Excel, PowerPoint) has been installed successfully! You may now use apps without any issue.', 'Office Download', 'OK', 'Information')
