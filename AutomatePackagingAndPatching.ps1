<#
.SYNOPSIS
    Auto-package Applications for Cloudpager.
.DESCRIPTION
    This script combines the Evergreen Module, Chocolatey, WinGet and Numecent Auto-package feature to automate application packaging of the latest version of public consumer versions of applications
    This script is published as an example and is wholly owned by Rory Monaghan. This is not a product or script supported by any vendor or firm. No guarantees are made or expressed. Any testing should be performed in a non-production environment.
.PARAMETER AppName
    Name of Application as found in the Evergreen or Chocolatey e.g. GoogleChrome. This can also be set to an application id in WinGet e.g. Google.Chrome.
.PARAMETER Publisher
    Name of the application's manufacturer.
.PARAMETER Sourcepackagetype
    Define the type of package you wish to auto-package e.g. msi, exe or msix. Other formats are not supported at this time.
.PARAMETER Sourcechannel
    If you wish to define a certain channel such as a stable channel, dev, beta etc. you can define with this parameter.
.PARAMETER Sourceplatform
    Some applications have different supported platforms such as a specific VDI version, if applicable this can be defined.
.PARAMETER Sourcelanguage
    Some applications listed in the Evergreen module have multiple language e.g. Adobe Acrobat Reader DC. In this case, you can select the lanugage that applies.
.PARAMETER image_file_path
    Provide the full path to an image for the application, preferably 512 x 512 in size.
.PARAMETER Arguments
    Passing install arguments for a silent install may be required for an exe installer. This is not required for msi or msix package types.
.PARAMETER CommandLine
    The CommandLine must be set to the full path of the main executable for the application.
.PARAMETER WorkdpodID
    If you wish to automatically publish the application to a Cloudpager Workdpod, pass the WorkpodID here e.g. You may have an early adopters Workpod for UAT.
.PARAMETER Description
    If you would like to set a description for the application, do this here.
.PARAMETER Chocolatey
    If you would like to use Chocolatey as the source for your application, set this parameter to $true otherwise set it to $false. Also ensure the AppName is set to the corresponding application name as it appears in Chocolatey.
.PARAMETER Evergreen
    If you would like to use the Evergeen PowerShell module as the source for your application, set this parameter to $true otherwise set it to $false. Also ensure the AppName is set to the corresponding application name as it appears in the module.
.PARAMETER WinGet
    If you would like to use WinGet as the source for your application, set this parameter to $true otherwise set it to $false. Also ensure the AppName is set to the corresponding application id as it appears in WinGet.
.REQUIRES PowerShell Version 5.0, Cloudpager, WinGet, Chocolatey and Evergreen PowerShell modules are required, the PSOpenAI module is optional. You will require this module if you wish to integrate with the OpenAI
    API. To test this capability, find sections of code that are commented out - this code can uncommented along with line for setting your OpenAI API Key to automatically set your description and publisher info. Read the comments in the code for more info. 
    You must have Cloudpaging Studio on the packaging VM along with Numecent's CreateJson.ps1 and studio-nip.ps1. You should run the CloudpagingStudio-prep.ps1 on your packaging VM before taking a snapshot. 
.EXAMPLE
    >AutomateEvergreenPackaging.ps1 -AppName "GoogleChrome" -publisher "Google" -sourcepackagetype "msi" -sourcechannel "stable" -image_file_path "\\ImageServer\Images\Chrome.png" -CommandLine "C:\Program Files\Google\Chrome\Application\chrome.exe" -WorkpodID "<id>" -Description "Google Chrome is the world's most popular web browser." -Chocolatety $false -Evergreen $true -WinGet $false
    >AutomateEvergreenPackaging.ps1 -AppName "NotepadPlusPlus" -publisher "Don Ho" -sourcepackagetype "exe" -sourceplatform "Windows" -image_file_path "\\ImageServer\Images\NotepadPlusPlus.png" -Arguments " /S" -CommandLine "C:\Program Files\Notepad++\notepad++.exe" -WorkpodID "<id>" -Chocolatety $false -Evergreen $false -WinGet $true
#>

Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$AppName,

   [Parameter(Mandatory=$True)]
   [string]$Publisher,
   
   [Parameter(Mandatory=$True)]
   [boolean]$Chocolatey,
   
   [Parameter(Mandatory=$True)]
   [boolean]$Evergreen,

   [Parameter(Mandatory=$True)]
   [boolean]$WinGet,

   [Parameter(Mandatory=$True)]
   [string]$Sourcepackagetype,

   [Parameter(Mandatory=$False)]
   [string]$Sourcechannel,

   [Parameter(Mandatory=$False)]
   [string]$Sourceplatform,

   [Parameter(Mandatory=$False)]
   [string]$Sourcelanguage,
   
   [Parameter(Mandatory=$True)]
   [string]$image_file_path,

   [Parameter(Mandatory=$False)]
   [string]$Arguments,

   [Parameter(Mandatory=$True)]
   [string]$CommandLine,

   [Parameter(Mandatory=$False)]
   [string]$WorkpodID,

   [Parameter(Mandatory=$True)]
   [string]$Description,

   [Parameter(Mandatory=$false)]
   [string[]]$registryexclusions,

   [Parameter(Mandatory=$false)]
   [string[]]$FileExclusion
)

#Enter values for these variables before using script
$skey = "<skey>"

#$TeamsURI = "<WebHookURI>"

#$AppsDashboardURL = '<Cloudpager Apps Dashboard URL>'

#Remove the comments for the next 2 lines and add OpenAI API Key to use Open AI API as part of the script
#Import-Module PSOpenAI
#$global:OPENAI_API_KEY = '<APIKey>'

if ($WinGet -eq $True -and $Evergreen -eq $True) {
    throw 'You must not set WinGet and Evergreen to True, please choose a single source.'
}

if ($WinGet -eq $True -and $Chocolatey -eq $True) {
    throw 'You must not set WinGet and Chocolatey to True, please choose a single source.'
}

if ($Evergreen -eq $True -and $Chocolatey -eq $True) {
    throw 'You must not set Chocolatey and Evergreen to True, please choose a single source.'
}

if ($WinGet -eq $True -and $Evergreen -eq $True -and $Chocolatey -eq $True) {
    throw 'You must not set WinGet, Evergreen and Chocolatey to True, please choose a single source.'
}

$CreateJSONFile = Test-Path ".\CreateJson.ps1" -PathType Leaf
if($CreateJSONFile -eq $False)
{
Write-Error "The CreateJson script is missing or you are running this script from a different directory. Ensure the CreateJson.ps1 script and all other scripts are placed in the scripts directory."
}

$studioNIPFile = Test-Path ".\studio-nip.ps1" -PathType Leaf
if($studioNIPFile -eq $False)
{
Write-Error "The studio-nip script is missing or you are running this script from a different directory. Ensure the studio-nip.ps1 script and all other scripts are placed in the scripts directory."
}
$webClient = New-Object System.Net.WebClient

$LocalJsonScript = Get-Content -Path ".\CreateJson.ps1"

$GitJsonScript = "https://raw.githubusercontent.com/Numecent/Automated-Packaging/Powershell-json-generation/Powershell-Generator/NIP_Software/Scripts/CreateJson.ps1"

if(Compare-Object  "C:\NIP_Software\Scripts\CreateJson.ps1"  ($GitJsonScript -replace '\r?\n\z' -split '\r?\n' ))

 {
Remove-Item "C:\NIP_Software\Scripts\CreateJson.ps1"
$webClient.DownloadFile($GitJsonScript, "C:\NIP_Software\Scripts\CreateJson.ps1")
 }

if($WinGet -eq $True -or $Chocolatey -eq $True)
{
if($WinGet -eq $True)
{
Try
{
    # Try something that could cause an error
    winget search "$AppName"
}
Catch
{
    # Catch any error
    Write-Error "An error has occurred retrieving data for $AppName from the Windows Package Manager. Ensure the App Installer app is installed, if errors continue try to run a query manually using winget search $AppName"
}

$NIPDirectory = "C:\NIP_Software"

[string]$Name
[string]$Id
[string]$Version

$searchResult = winget search $AppName | Out-String

$lines = $searchResult -split "`r`n"

# Find the line that starts with Name, it contains the header
$fl = 0
while (-not $lines[$fl].StartsWith("Name"))
{
    $fl++
}

# Line $fl has the header, we can find char where we find ID and Version
$NameStart = $lines[$fl].IndexOf("Name")
$idStart = $lines[$fl].IndexOf("Id")
$versionStart = $lines[$fl].IndexOf("Version")
$sourceStart = $lines[$fl].IndexOf("Source")

# Now cycle in real package and split accordingly
$searchList = @()
$found = $false
For ($i = $fl + 1; $i -lt $lines.Length; $i++) 
{
    $line = $lines[$i]
    if ($line.Length -gt ($sourceStart + 1) -and -not $line.StartsWith('-'))
    {
        $name = $line.Substring(0, $idStart).TrimEnd()
        $id = $line.Substring($idStart, $versionStart - $idStart).TrimEnd()
        $version = $line.Substring($versionStart, $sourceStart - $versionStart).TrimEnd()
        
        # If this is the app we are looking for, stop processing further
        if ($id -eq $AppName)
        {
            $found = $true
            break
        }
    }
}

if($found)
{
    Write-Output "$AppName found in WinGet! Name: $name, Id: $id, Version: $version"
}
else
{
    Write-Error "$AppName not found!"
}


Try
{
    # Try something that could cause an error
    Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion
}
Catch
{
    # Catch any error
    Write-Error "An error has occurred retrieving data for $Name from Cloudpager. Ensure the Cloudpager API is installed."
}

$Curversion = Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion

$Curversion = $Curversion | measure -Maximum | select -ExpandProperty Maximum

If($Version -ne $Curversion -or $Curversion -eq $null){

New-Item "$NIPDirectory\Auto\Install.cmd"

Set-Content "$NIPDirectory\Auto\Install.cmd" "winget install $AppName"

.\CreateJson.ps1 -Filepath "$NIPDirectory\Auto\Install.cmd" -Description $Description -Name $Name -Arguments " " -StudioCommandLine $CommandLine -outputfolder "$NIPDirectory\Auto" -iconFile $CommandLine

$config_file_path = Get-ChildItem -Path "$NIPDirectory\Auto" -Filter *.json | ForEach-Object{$_.FullName}

.\studio-nip.ps1 -config_file_path $config_file_path

$PackageFile = Get-ChildItem -Path "$NIPDirectory\Auto" -Filter *.stp | ForEach-Object{$_.FullName}

Add-CloudpagerApplication -SubscriptionKey $skey -Filepath $PackageFile -Name $Name -AppVersion $Version -Publisher $publisher -ImagePath $image_file_path -Description $Description -PublishComment "Uploaded using API" -Force 

If($WorkpodID){
Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$Name" -PublishComment "Added $Name $Version" -Confirm -Force
}
}
else
{
Write-Output "$Name is already published in your Cloudpager tenant."
}
}

if($Chocolatey -eq $True)
{
Try
{
    # Try something that could cause an error
    choco search $AppName
}
Catch
{
    # Catch any error
    Write-Error "An error has occurred retrieving data for $AppName from the Windows Package Manager. Ensure the App Installer app is installed, if errors continue try to run a query manually using winget search $AppName"
}


Try
{
    # Try something that could cause an error
    Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion
}
Catch
{
    # Catch any error
    Write-Error "An error has occurred retrieving data for $Name from Cloudpager. Ensure the Cloudpager API is installed."
}


$Curversion = Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion

$Curversion = $Curversion | measure -Maximum | select -ExpandProperty Maximum

$NIPDirectory = "C:\NIP_Software"

$AppInfo = choco info $AppName | Out-File "$NIPDirectory\Output\ChocoInfo.json" | ConvertTo-Json

$Summary = (Get-Content $NIPDirectory\Output\ChocoInfo.json) -match 'Summary'

$Summary = $Summary -replace " Summary: "

$Version = (Get-Content $NIPDirectory\Output\ChocoInfo.json) -match "$AppName"

$Version = $Version -replace "$AppName "

$Version = $Version.split(' ')[0]

$Title = (Get-Content $NIPDirectory\Output\ChocoInfo.json) -match 'Title'

$Title = $Title -replace " Title: "
$Title = $Title.split('|')[0]
$Title = $Title.trim()

$Name = $Title

If($LatestVersion -ne $Curversion -or $Curversion -eq $null){

New-Item "$NIPDirectory\Auto\Install.cmd"

Set-Content "$NIPDirectory\Auto\Install.cmd" "choco install $AppName -y"

.\CreateJson.ps1 -Filepath "$NIPDirectory\Auto\Install.cmd" -Description $Description -Name $Name -Arguments " " -StudioCommandLine $CommandLine -outputfolder "$NIPDirectory\Auto"

$config_file_path = Get-ChildItem -Path "$NIPDirectory\Auto" -Filter *.json | ForEach-Object{$_.FullName}

.\studio-nip.ps1 -config_file_path $config_file_path

$PackageFile = Get-ChildItem -Path "$NIPDirectory\Auto" -Filter *.stp | ForEach-Object{$_.FullName}

Add-CloudpagerApplication -SubscriptionKey $skey -Filepath $PackageFile -Name $Name -AppVersion $Version -Publisher $publisher -ImagePath $image_file_path -Description $Description -PublishComment "Uploaded using API" -Force
}
}
}
else
{

Try
{
    # Try something that could cause an error
    Find-EvergreenApp -Name $AppName | Where-Object { ($_.Name -eq $AppName) } | Select -ExpandProperty Application | Sort-Object { [System.Math]::Abs([System.String]::Compare($_, $AppName)) } | Select-Object -First 1
}
Catch
{
    # Catch any error
    Write-Host "An error has occurred retrieving data for $AppName from the Evergreen PowerShell Module. Ensure the module is loaded, if errors continue try to run a query manually using Find-EvergreenApp -Name $AppName"
}


$FriendlyName = Find-EvergreenApp -Name $AppName | Where-Object { ($_.Name -eq $AppName) } | Select -ExpandProperty Application | Sort-Object { [System.Math]::Abs([System.String]::Compare($_, $AppName)) } | Select-Object -First 1

#Remove comment for the line below and change Publisher parameter to Mandatory=$False to let OpenAI API populate the Publisher for you.
#$PublisherShort = Request-ChatCompletion -Message "Tell me what company makes $FriendlyName, only return the name of the company"
#$Publisher = $PublisherShort.Answer

Try
{
    # Try something that could cause an error
    Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $FriendlyName} | Select -ExpandProperty AppVersion
}
Catch
{
    # Catch any error
    Write-Host "An error has occurred retrieving data for $FriendlyName from the Cloudpager PowerShell Module. Ensure the module is loaded, if errors continue try to run a query manually using Get-CloudpagerApplication -SubscriptionKey $skey -Name $FriendlyName with double quotes around the app name."
}

$Curversion = Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $FriendlyName} | Select -ExpandProperty AppVersion

$Curversion = $Curversion | measure -Maximum | select -ExpandProperty Maximum

$AppCheck = Get-EvergreenApp -Name "$AppName"

if ($AppCheck.Count -eq 1) {
    $DownloadURL = Get-EvergreenApp -Name $AppName | Select -ExpandProperty URI
    $LatestVersion = Get-EvergreenApp -Name $AppName | Select -ExpandProperty Version
}
else
{

$BaseTest = Get-EvergreenApp -Name $AppName

$LatestVersion = Get-EvergreenApp -Name $AppName | Where-Object { if (!$_.Architecture -or $_.Architecture -eq "x64") {$true} else {$false} } | Where-Object { if (!$_.Channel -or $_.Channel -eq $sourcechannel) {$true} else {$false} } | Where-Object { if (!$_.Type -or $_.Type -eq $sourcepackagetype) {$true} else {$false} } | Where-Object { if (!$_.Platform -or $_.Platform -eq $sourceplatform) {$true} else {$false} } | Where-Object { if (!$_.Language -or $_.Language -eq $sourcelanguage) {$true} else {$false} } | Select -ExpandProperty Version | Select-Object -First 1
$DownloadURL = Get-EvergreenApp -Name $AppName | Where-Object { if (!$_.Architecture -or $_.Architecture -eq "x64") {$true} else {$false} } | Where-Object { if (!$_.Channel -or $_.Channel -eq $sourcechannel) {$true} else {$false} } | Where-Object { if (!$_.Type -or $_.Type -eq $sourcepackagetype) {$true} else {$false} } | Where-Object { if (!$_.Platform -or $_.Platform -eq $sourceplatform) {$true} else {$false} } | Where-Object { if (!$_.Language -or $_.Language -eq $sourcelanguage) {$true} else {$false} } | Select -ExpandProperty URI | Select-Object -First 1
}

$ProjectFolder = "C:\NIP_Software\$AppName"

$DownloadFilePath = "C:\NIP_Software\Auto\Latest$AppName.$sourcepackagetype"

If($LatestVersion -ne $Curversion -or $Curversion -eq $null){

Write-Output "New version detected. Now auto-packaging!"

If(!(test-path $ProjectFolder))
{
      New-Item -ItemType Directory -Force -Path $ProjectFolder
      New-Item -ItemType Directory -Force -Path "$ProjectFolder\Source"
      New-Item -ItemType Directory -Force -Path "$ProjectFolder\Output"
}

$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($DownloadURL, $DownloadFilePath)

#Invoke-WebRequest -Uri $DownloadURL -OutFile $DownloadFilePath

If($sourcepackagetype -eq "msix")
{
$PackageFile = Get-ChildItem -Path "C:\NIP_Software\Auto" -Filter *.msix | ForEach-Object{$_.FullName}
Add-CloudpagerApplication -SubscriptionKey $skey -Filepath $PackageFile -Name $FriendlyName -AppVersion $LatestVersion -Publisher $publisher -ImagePath $image_file_path -Description $Description -PublishComment "Uploaded using API" -Force 
If($WorkpodID){
Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$FriendlyName" -PublishComment "Added $FriendlyName $LatestVersion" -Confirm -Force
}

}
else
{

#Remove comment for the line below and change Description parameter to let OpenAI API populate the Publisher for you.
#$DescriptionShort = Request-ChatCompletion -Message "Describe $Publisher $FriendlyName in 30 words or less"
#$Description = $DescriptionShort.Answer

.\CreateJson.ps1 -Filepath $DownloadFilePath -Description $Description -Name $FriendlyName -Arguments $Arguments -RegistryExclusions $registryexclusions -FileExclusions $fileexclusion -StudioCommandLine $CommandLine -outputfolder "$ProjectFolder\Output"

$config_file_path = Get-ChildItem -Path "C:\NIP_Software\Auto" -Filter *.json | ForEach-Object{$_.FullName}

.\studio-nip.ps1 -config_file_path $config_file_path

$PackageFile = Get-ChildItem -Path "$ProjectFolder\Output" -Filter *.stp | ForEach-Object{$_.FullName}

Add-CloudpagerApplication -SubscriptionKey $skey -Filepath $PackageFile -Name $FriendlyName -AppVersion $LatestVersion -Publisher $publisher -ImagePath $image_file_path -Description $Description -PublishComment "Uploaded using API" -Force 

Write-Output "$AppName $LatestVersion is now available in Cloudpager!"

#Remove comment for the line below to send a Teams Notification.
#.\SendTeamsMessage.ps1 -WebhookUri $TeamsURI -Title 'App Update' -Message "$AppName $LatestVersion is now available in Cloudpager!" -Proxy 'DoNotUse' -ButtonText 'View' -ButtonURI $AppsDashboardURL 


If($WorkpodID -and $WinGet -ne $True){
Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$FriendlyName" -PublishComment "Added $FriendlyName $LatestVersion" -Confirm -Force
}
}
}
else
{
Write-Output "Latest version of $AppName is already published in Cloudpager"
}
}
