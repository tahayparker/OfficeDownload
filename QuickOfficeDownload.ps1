# Download Office - Word, Excel, Powerpoint

# Make a temp dir
mkdir C:\temp

# Change to the temp dir
cd C:\temp

# Download the setup file
Invoke-WebRequest https://raw.githubusercontent.com/tahayparker/OfficeDownload/refs/heads/main/setup.exe -Output setup.exe

# Warn user to not open the apps
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show('Office will now be installed. Please do not open any apps until another message box appears', 'Office Download', 'OK', 'Warning')

# Run the setup file
Start-Process -FilePath C:\temp\setup.exe -ArgumentList "/configure https://raw.githubusercontent.com/tahayparker/OfficeDownload/refs/heads/main/TPPublic.xml" -Wait

# Remove the temp dir
Remove-Item C:\temp -Recurse -Force

# Run get.activated.win script to activate Office
& ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID /S

# Inform user via msgbox alert
[System.Windows.MessageBox]::Show('Office (Word, Excel, PowerPoint) has been installed successfully! You may now use apps without any issue.', 'Office Download', 'OK', 'Information')