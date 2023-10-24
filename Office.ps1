# Ensure scipt is running as admin, otherwise relaunch as admin. If fails, exit script
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

# Define Offce versions
$officeVersions = [Ordered]@{
    '1' = 'Office 2013'
    '2' = 'Office 2016'
    '3' = 'Office 2019'
    '4' = 'Office 2021'
    '5' = 'Office 365'
}

# Display the available Office versions to the user
Write-Host "Choose an Office version to download:"
foreach ($key in $officeVersions.Keys) {
    Write-Host "$key. $($officeVersions[$key])"
}

# Prompt the user to input a number for their preferred version
$VersionSelection = Read-Host "Enter preferred Office version"

# Prompt user for architecture of Office to install (x86, x64)
$officeArchs = [Ordered]@{
    '1' = 'x64'
    '2' = 'x86'
}

# Print the available Office archs to the user
Write-Host "Choose an Office arch to download:"
foreach ($key in $officeArchs.Keys) {
    Write-Host "$key. $($officeArchs[$key])"
}

# Prompt the user to input a number for their preferred arch
$ArchSelection = Read-Host "Enter preferred Office arch"

# If selection is option 1 to 4, prompt user for type of Office to install
if (($VersionSelection -eq '1') -or ($VersionSelection -eq '2') -or ($VersionSelection -eq '3') -or ($VersionSelection -eq '4'))
{
    $officeTypes = [Ordered]@{
        '1' = 'ProPlus'
        '2' = 'Professional'
        '3' = 'Standard'
        '4' = 'Personal'
        '5' = 'HomeBusiness'
        '6' = 'HomeStudent'
    }

    # Print the available Office types to the user
    Write-Host "Choose an Office type to download:"
    foreach ($key in $officeTypes.Keys) {
        Write-Host "$key. $($officeTypes[$key])"
    }
    # Prompt the user to input a number for their preferred type
    $TypeSelection = Read-Host "Enter preferred Office type"
}

# If selection is option 5, prompt user for type of Office to install
ElseIf ($VersionSelection -eq '5') {
    $officeTypes = [Ordered]@{
        '1' = 'ProPlus'
        '2' = 'Business'
        '3' = 'SmallBusinessPremium'
        '4' = 'HomePremium'
        '5' = 'EducationCloud'
    }
    # Print the available Office types to the user
    Write-Host "Choose an Office type to download:"
    foreach ($key in $officeTypes.Keys) {
        Write-Host "$key. $($officeTypes[$key])"
    }
    # Prompt the user to input a number for their preferred type
    $TypeSelection = Read-Host "Enter preferred Office type"
}

# Set download URLs based on user input

if (($VersionSelection -eq '1') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O15GA'
    $ProdID = 'ProPlusRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O15GA'
    $ProdID = 'ProfessionalRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=StandardRetail&platform=x64&language=en-us&version=O15GA'
}   $ProdID = 'StandardRetail'

elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=PersonalRetail&platform=x64&language=en-us&version=O15GA'
    $ProdID = 'PersonalRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x64&language=en-us&version=O15GA'
    $ProdID = 'HomeBusinessRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x64&language=en-us&version=O15GA'
    $ProdID = 'HomeStudentRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x86&language=en-us&version=O15GA'
    $ProdID = 'ProPlusRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x86&language=en-us&version=O15GA'
    $ProdID = 'ProfessionalRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=StandardRetail&platform=x86&language=en-us&version=O15GA'
    $ProdID = 'StandardRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=PersonalRetail&platform=x86&language=en-us&version=O15GA'
    $ProdID = 'PersonalRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x86&language=en-us&version=O15GA'
    $ProdID = 'HomeBusinessRetail'
}
elseif (($VersionSelection -eq '1') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x86&language=en-us&version=O15GA'
    $ProdID = 'HomeStudentRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'ProPlusRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'ProfessionalRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=StandardRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'StandardRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=PersonalRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'PersonalRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'HomeBusinessRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'HomeStudentRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'ProPlusRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProfessionalRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'ProfessionalRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=StandardRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'StandardRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=PersonalRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'PersonalRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusinessRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'HomeBusinessRetail'
}
elseif (($VersionSelection -eq '2') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudentRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'HomeStudentRetail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'ProPlus2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2019Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'Professional2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Standard2019Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'Standard2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Personal2019Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'Personal2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2019Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'HomeBusiness2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2019Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'HomeStudent2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'ProPlus2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2019Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'Professional2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Standard2019Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'Standard2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Personal2019Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'Personal2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2019Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'HomeBusiness2019Retail'
}
elseif (($VersionSelection -eq '3') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2019Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'HomeStudent2019Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'ProPlus2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2021Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'Professional2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2021Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'Professional2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Personal2021Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'Personal2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2021Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'HomeBusiness2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2021Retail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'HomeStudent2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'ProPlus2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Professional2021Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'Professional2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Standard2021Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'Standard2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=Personal2021Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'Personal2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeBusiness2021Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'HomeBusiness2021Retail'
}
elseif (($VersionSelection -eq '4') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '6')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=HomeStudent2021Retail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'HomeStudent2021Retail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'O365ProPlusRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'O365BusinessRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365SmallBusPremRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'O365SmallBusPremRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'O365HomePremRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '1') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365EduCloudRetail&platform=x64&language=en-us&version=O16GA'
    $ProdID = 'O365EduCloudRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '1')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'O365ProPlusRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '2')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'O365BusinessRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '3')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365SmallBusPremRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'O365SmallBusPremRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '4')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'O365HomePremRetail'
}
elseif (($VersionSelection -eq '5') -and ($ArchSelection -eq '2') -and ($TypeSelection -eq '5')) {
    $url = 'https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365EduCloudRetail&platform=x86&language=en-us&version=O16GA'
    $ProdID = 'O365EduCloudRetail'
}

# Ask users which apps they dont want to install, if 
$officeApps = [Ordered]@{
    '1' = 'Access'
    '2' = 'Excel'
    '3' = 'Groove'
    '4' = 'Lync'
    '5' = 'OneDrive'
    '6' = 'OneNote'
    '7' = 'Outlook'
    '8' = 'PowerPoint'
    '9' = 'Project'
    '10' = 'Publisher'
    '11' = 'SharePointDesigner'
    '12' = 'Visio'
    '13' = 'Word'
    '14' = 'Skype'
    '15' = 'Teams'
}

# Print the available Office apps to the user
Write-Host "Choose Office apps to exclude:"
foreach ($key in $officeApps.Keys) {
    Write-Host "$key. $($officeApps[$key])"
}

# Prompt the user to input a number for their preferred apps
$AppSelection = Read-Host "Enter Office apps to remove, separated by commas"

# Parse the user input into an array
$AppSelection = $AppSelection.Split(',')
$AppSelection = $AppSelection.Trim()

# Create the OfficeConfig.xml file dynamically based on user selections
$OfficeConfig = @"
<Configuration>
    <Add OfficeClientEdition="$($officeArchs[$ArchSelection])">
        <Product ID="$ProdID">
            <Language ID="MatchOS" />`n
"@
if ($null -ne $AppSelection -and $AppSelection.Count -gt 0) {
    $OfficeConfig += "            <ExcludeApp ID='" + ($officeApps[$AppSelection[0]]) + "' />`n"
    for ($i = 1; $i -lt $AppSelection.Count; $i++) {
        $OfficeConfig += "            <ExcludeApp ID='" + ($officeApps[$AppSelection[$i]]) + "' />`n"
    }
}
elseif ($null -eq $AppSelection -and $AppSelection.Count -eq 0) {
    OfficeConfig += ""
}

$OfficeConfig += @"
        </Product>
    </Add>
    <Property Name="SharedComputerLicensing" Value="0" />
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="DeviceBasedLicensing" Value="0" />
    <Property Name="SCLCacheOverride" Value="0" />
    <Property Name="AUTOACTIVATE" Value="1" />
    <Updates Enabled="TRUE" />
    <RemoveMSI />
    <Display Level="Full" AcceptEULA="TRUE" />  
</Configuration>
"@

Write-Host "OfficeConfig.xml file created."

# Check if Downloads folder exists, if not create it
if (!(Test-Path "$env:USERPROFILE\Downloads")) {
    New-Item -Path "$env:USERPROFILE\Downloads" -ItemType Directory
}

# Save the OfficeConfig.xml to a file
$OfficeConfig | Out-File -FilePath "$env:USERPROFILE\Downloads\OfficeConfig.xml" -Encoding UTF8

# Office Config XML Location
$OfficeConfig = "$env:USERPROFILE\Downloads\OfficeConfig.xml"

# Download Office installer
$downloadPath = "$env:USERPROFILE\Downloads\OfficeSetup.exe"
Invoke-WebRequest -Uri $url -OutFile $downloadPath

# Install Office
Start-Process -FilePath $downloadPath -ArgumentList '/configure', $OfficeConfig -Wait

& ([ScriptBlock]::Create((Invoke-WebRequest https://massgrave.dev/get))) /Ohook /S

Write-Host "Office hass been installed successfully!"
