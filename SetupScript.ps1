<#
.SYNOPSIS
    Creates auto-packaging directories on packaging VM and downloads scripts.
.DESCRIPTION
    This script will create the packaging directories for you and download scripts.
#>

$webClient = New-Object System.Net.WebClient

$NIPDirectory = "C:\NIP_Software"
$NIPScriptsDirectory = "$NIPDirectory\Scripts"

$PrepScript = "https://raw.githubusercontent.com/Numecent/Automated-Packaging/master/CloudpagingStudio-prep.ps1"
$NIPScript = "https://raw.githubusercontent.com/Numecent/Automated-Packaging/master/studio-nip.ps1"
$JSONScript = "https://raw.githubusercontent.com/Numecent/Automated-Packaging/Powershell-json-generation/Powershell-Generator/NIP_Software/Scripts/CreateJson.ps1"
$AutoInstallScript = 'https://raw.githubusercontent.com/Numecent/Automated-Packaging/Powershell-json-generation/Powershell-Generator/NIP_Software/Scripts/Auto_Install_from_json.ps1'
$AutopackageScript = "https://raw.githubusercontent.com/Rorymon/AutomateEvergreenPackaging/main/AutomateEvergreenPackaging.ps1"

If(!(test-path $NIPDirectory))
{
      New-Item -ItemType Directory -Force -Path $NIPDirectory
      New-Item -ItemType Directory -Force -Path "$NIPDirectory\Auto"
      New-Item -ItemType Directory -Force -Path "$NIPDirectory\Output"
      New-Item -ItemType Directory -Force -Path $NIPScriptsDirectory
}

$webClient.DownloadFile($AutoInstallScript, "$NIPScriptsDirectory\Auto_Install_from_json.ps1")
$webClient.DownloadFile($PrepScript, "$NIPScriptsDirectory\CloudpagingStudio-prep.ps1")
$webClient.DownloadFile($NIPScript, "$NIPScriptsDirectory\studio-nip.ps1")
$webClient.DownloadFile($JSONScript, "$NIPScriptsDirectory\CreateJson.ps1")
$webClient.DownloadFile($AutopackageScript, "$NIPScriptsDirectory\AutomateEvergreenPackaging.ps1")
