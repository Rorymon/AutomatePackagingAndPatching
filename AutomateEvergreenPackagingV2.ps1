<#
.SYNOPSIS
    Auto-package Applications for Cloudpager.
.DESCRIPTION
    This script combines the Evergreen Module, Nevergreen Module and Numecent Auto-package feature to automate application
    packaging of the latest version of public consumer versions of applications.

    In addition to the Numecent Non-Interactive Packager (NIP) sources (Evergreen, WinGet, Chocolatey), the script now
    supports Cloudpager AI Packaging, which builds the package server-side (no NIP / Cloudpaging Studio required):
      * -AIPackagingWinGet : AI-package directly from the Cloudpager WinGet catalog (uses -AppName as the PackageIdentifier).
      * -AIPackagingMsi     : AI-package from a supplied MSI (-MsiPath). ProductVersion is auto-extracted from the MSI.
.PARAMETER AppName
    Name of Application as found in the Evergreen or Nevergreen Find- functions.
    For -AIPackagingWinGet this is the WinGet PackageIdentifier (e.g. "7zip.7zip").
    For -AIPackagingMsi this is the display Name used in Cloudpager.
.PARAMETER Publisher
    Name of the application's manufacturer. Required for Evergreen/WinGet/Chocolatey.
    Optional for AI Packaging: if omitted for -AIPackagingWinGet the catalog value is used; for -AIPackagingMsi the MSI
    Manufacturer property is used as a fallback.
.PARAMETER Chocolatey
    Set to $True to source the application from Chocolatey (NIP packaging).
.PARAMETER Evergreen
    Set to $True to source the application from Evergreen (NIP packaging).
.PARAMETER WinGet
    Set to $True to source the application from the Windows Package Manager (NIP packaging).
.PARAMETER AIPackagingWinGet
    Set to $True to use Cloudpager AI Packaging against the Cloudpager WinGet catalog. -AppName is the PackageIdentifier.
.PARAMETER AIPackagingMsi
    Set to $True to use Cloudpager AI Packaging with a supplied MSI. Provide the MSI via -MsiPath.
.PARAMETER MsiPath
    Full path to the MSI to AI-package (only used with -AIPackagingMsi).
.PARAMETER AppVersion
    Optional explicit version. Overrides the auto-detected version for the AI Packaging modes.
.PARAMETER Sourcepackagetype
    Define the type of package you wish to auto-package e.g. msi, exe or msix. Required for -Evergreen only.
.PARAMETER Sourcechannel
    If you wish to define a certain channel such as a stable channel, dev, beta etc. you can define with this parameter.
.PARAMETER Sourceplatform
    Some applications have different supported platforms such as a specific VDI version, if applicable this can be defined.
.PARAMETER Sourcelanguage
    Some applications listed in the Evergreen module have multiple languages e.g. Adobe Acrobat Reader DC.
.PARAMETER image_file_path
    Provide the full path OR an http(s) URL to an image for the application, preferably 512 x 512 in size.
    If a URL is supplied it is downloaded locally and the local copy is used. Required for Evergreen/WinGet/Chocolatey;
    optional (but recommended) for AI Packaging.
.PARAMETER Arguments
    Passing install arguments for a silent install may be required for an exe installer. For -AIPackagingMsi these are
    passed through as -InstallerArguments.
.PARAMETER CommandLine
    The CommandLine must be set to the full path of the main executable for the application. Required for NIP modes only.
.PARAMETER WorkpodID
    If you wish to automatically publish the application to a Cloudpager Workpod, pass the WorkpodID here.
.PARAMETER Description
    Description for the application. Required for ALL modes (AI Packaging auto-fill descriptions are generic and can bleed
    through to Citrix Studio/StoreFront, Intune and Cloudpager storefronts, so a real description is enforced).
.REQUIRES PowerShell Version 5.0, Cloudpager PowerShell module. Evergreen/NIP requirements apply only to the NIP sources.
.EXAMPLE
    # Evergreen (NIP)
    >AutomateEvergreenPackaging.ps1 -AppName "GoogleChrome" -Evergreen $True -Publisher "Google" -Sourcepackagetype "msi" -Sourcechannel "stable" -image_file_path "https://raw.githubusercontent.com/Rorymon/icons/main/icons/GoogleChrome.png" -CommandLine "C:\Program Files\Google\Chrome\Application\chrome.exe" -Description "Google Chrome is the world's most popular web browser."
.EXAMPLE
    # Chocolatey (NIP)
    >AutomateEvergreenPackaging.ps1 -AppName "7zip" -Publisher "Igor Pavlov" -Chocolatey $True -Description "7-Zip is an open source utility that supports various compression file formats." -image_file_path "https://raw.githubusercontent.com/Rorymon/icons/main/icons/7-Zip.png" -CommandLine "C:\Program Files\7-Zip\7zFM.exe"
.EXAMPLE
    # WinGet (NIP)
    >AutomateEvergreenPackaging.ps1 -AppName "7zip.7zip" -Publisher "Igor Pavlov" -WinGet $True -Description "7-Zip is an open source utility that supports various compression file formats." -image_file_path "https://raw.githubusercontent.com/Rorymon/icons/main/icons/7-Zip.png" -CommandLine "C:\Program Files\7-Zip\7zFM.exe"
.EXAMPLE
    # AI Packaging from the WinGet catalog
    >AutomateEvergreenPackaging.ps1 -AppName "7zip.7zip" -AIPackagingWinGet $True -Description "7-Zip is an open source utility that supports various compression file formats." -image_file_path "https://raw.githubusercontent.com/Rorymon/icons/main/icons/7-Zip.png"
.EXAMPLE
    # AI Packaging from a supplied MSI
    >AutomateEvergreenPackaging.ps1 -AppName "Zoom" -AIPackagingMsi $True -MsiPath "C:\ProgramData\ForAIPackaging\ZoomInstallerFull.msi" -Publisher "Zoom" -Description "Zoom is a popular web conferencing and communication platform." -Arguments "ZConfig=nofacebook=1;nogoogle=1 /qn" -image_file_path "https://raw.githubusercontent.com/Rorymon/icons/main/icons/Zoom.png"
#>

Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$AppName,

   [Parameter(Mandatory=$False)]
   [string]$Publisher,

   [Parameter(Mandatory=$False)]
   [boolean]$Chocolatey,

   [Parameter(Mandatory=$False)]
   [boolean]$Evergreen,

   [Parameter(Mandatory=$False)]
   [boolean]$WinGet,

   [Parameter(Mandatory=$False)]
   [boolean]$AIPackagingWinGet,

   [Parameter(Mandatory=$False)]
   [boolean]$AIPackagingMsi,

   [Parameter(Mandatory=$False)]
   [string]$MsiPath,

   [Parameter(Mandatory=$False)]
   [string]$AppVersion,

   [Parameter(Mandatory=$False)]
   [string]$Sourcepackagetype,

   [Parameter(Mandatory=$False)]
   [string]$Sourcechannel,

   [Parameter(Mandatory=$False)]
   [string]$Sourceplatform,

   [Parameter(Mandatory=$False)]
   [string]$Sourcelanguage,

   [Parameter(Mandatory=$False)]
   [string]$image_file_path,

   [Parameter(Mandatory=$False)]
   [string]$Arguments,

   [Parameter(Mandatory=$False)]
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

# ------------------------------------------------------------------------------------
# Helper functions
# ------------------------------------------------------------------------------------

function Resolve-ImagePath {
    <#
      If the supplied value is an http(s) URL, download it to the local icon cache and
      return the local path. Otherwise return the value unchanged.
    #>
    param([string]$PathOrUrl)

    if ([string]::IsNullOrWhiteSpace($PathOrUrl)) { return $null }

    if ($PathOrUrl -match '^(?i)https?://') {
        $iconDir = "C:\NIP_Software\Icons"
        if (!(Test-Path $iconDir)) { New-Item -ItemType Directory -Force -Path $iconDir | Out-Null }

        $fileName = [System.IO.Path]::GetFileName(([Uri]$PathOrUrl).AbsolutePath)
        if ([string]::IsNullOrWhiteSpace($fileName)) {
            $fileName = "icon_$(Get-Date -Format 'yyyyMMddHHmmss').png"
        }

        $dest = Join-Path $iconDir $fileName
        Write-Host "Downloading image from $PathOrUrl to $dest"
        (New-Object System.Net.WebClient).DownloadFile($PathOrUrl, $dest)
        return $dest
    }

    return $PathOrUrl
}

function Get-MsiProperty {
    <#
      Reads a Property table value (e.g. ProductVersion, Manufacturer, ProductName) from an MSI.
    #>
    param(
        [Parameter(Mandatory=$True)][string]$Path,
        [Parameter(Mandatory=$True)][string]$Property
    )

    $installer = $null
    $database  = $null
    $view      = $null
    $record    = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database  = $installer.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$installer,@($Path,0))
        $query     = "SELECT Value FROM Property WHERE Property = '$Property'"
        $view      = $database.GetType().InvokeMember('OpenView','InvokeMethod',$null,$database,@($query))
        $view.GetType().InvokeMember('Execute','InvokeMethod',$null,$view,$null) | Out-Null
        $record    = $view.GetType().InvokeMember('Fetch','InvokeMethod',$null,$view,$null)
        if ($null -ne $record) {
            return $record.GetType().InvokeMember('StringData','GetProperty',$null,$record,1)
        }
        return $null
    }
    finally {
        foreach ($obj in @($record,$view,$database,$installer)) {
            if ($null -ne $obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
        }
    }
}

function Get-WinGetPackageInfo {
    <#
      Parses "winget search <Identifier>" output and returns Name/Id/Version for an exact Id match.
    #>
    param([Parameter(Mandatory=$True)][string]$Identifier)

    $searchResult = winget search $Identifier | Out-String
    $lines = $searchResult -split "`r`n"

    $fl = 0
    while ($fl -lt $lines.Length -and -not $lines[$fl].StartsWith("Name")) { $fl++ }
    if ($fl -ge $lines.Length) { return $null }

    $idStart      = $lines[$fl].IndexOf("Id")
    $versionStart = $lines[$fl].IndexOf("Version")
    $sourceStart  = $lines[$fl].IndexOf("Source")

    for ($i = $fl + 1; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        if ($line.Length -gt ($sourceStart + 1) -and -not $line.StartsWith('-')) {
            $n  = $line.Substring(0, $idStart).TrimEnd()
            $id = $line.Substring($idStart, $versionStart - $idStart).TrimEnd()
            $v  = $line.Substring($versionStart, $sourceStart - $versionStart).TrimEnd()
            if ($id -eq $Identifier) {
                return [pscustomobject]@{ Name = $n; Id = $id; Version = $v }
            }
        }
    }
    return $null
}

function Send-TeamsAdaptiveCard {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$WebhookUrl,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ImageUrl,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$LinkUrl,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ButtonText = 'View'
    )

    $cardBody = @(
        @{
            type   = 'TextBlock'
            text   = $Title
            weight = 'Bolder'
            size   = 'Large'
            wrap   = $true
        }
        @{
            type = 'TextBlock'
            text = $Message
            wrap = $true
        }
    )

    # Add an image only when an image URL was supplied.
    if ($ImageUrl) {
        $cardBody += @{
            type    = 'Image'
            url     = $ImageUrl
            size    = 'Auto'
            altText = $Title
        }
    }

    $card = @{
        '$schema' = 'https://adaptivecards.io/schemas/adaptive-card.json'
        type      = 'AdaptiveCard'
        version   = '1.4'
        body      = $cardBody
    }

    # Add a clickable button only when a destination URL was supplied.
    if ($LinkUrl) {
        $card.actions = @(
            @{
                type  = 'Action.OpenUrl'
                title = $ButtonText
                url   = $LinkUrl
            }
        )
    }

    $json = $card | ConvertTo-Json -Depth 20

    try {
        $response = Invoke-WebRequest `
            -Uri $WebhookUrl `
            -Method Post `
            -ContentType 'application/json; charset=utf-8' `
            -Body $json `
            -ErrorAction Stop

        if ($response.StatusCode -notin 200, 202) {
            throw "Teams returned HTTP status $($response.StatusCode)."
        }

        Write-Verbose "Adaptive Card submitted to Teams. HTTP $($response.StatusCode)."
    }
    catch {
        throw "Failed to send Teams Adaptive Card: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------------------------
# Source selection & validation
# ------------------------------------------------------------------------------------

$sources = @()
if ($Evergreen)         { $sources += 'Evergreen' }
if ($WinGet)            { $sources += 'WinGet' }
if ($Chocolatey)        { $sources += 'Chocolatey' }
if ($AIPackagingWinGet) { $sources += 'AIPackagingWinGet' }
if ($AIPackagingMsi)    { $sources += 'AIPackagingMsi' }

if ($sources.Count -eq 0) {
    throw 'Select exactly one source ($True): -Evergreen, -WinGet, -Chocolatey, -AIPackagingWinGet, or -AIPackagingMsi.'
}
if ($sources.Count -gt 1) {
    throw "Select only ONE source. You selected: $($sources -join ', ')."
}
$Source = $sources[0]

# Per-mode required-field validation (fields are no longer Mandatory at the Param level
# because they do not all apply to the AI Packaging modes).
$missing = @()
switch ($Source) {
    'Evergreen' {
        if (-not $Publisher)         { $missing += 'Publisher' }
        if (-not $Sourcepackagetype) { $missing += 'Sourcepackagetype' }
        if (-not $image_file_path)   { $missing += 'image_file_path' }
        if (-not $CommandLine)       { $missing += 'CommandLine' }
    }
    'WinGet' {
        if (-not $Publisher)       { $missing += 'Publisher' }
        if (-not $image_file_path) { $missing += 'image_file_path' }
        if (-not $CommandLine)     { $missing += 'CommandLine' }
    }
    'Chocolatey' {
        if (-not $Publisher)       { $missing += 'Publisher' }
        if (-not $image_file_path) { $missing += 'image_file_path' }
        if (-not $CommandLine)     { $missing += 'CommandLine' }
    }
    'AIPackagingMsi' {
        if (-not $MsiPath) {
            $missing += 'MsiPath'
        }
        elseif (-not (Test-Path $MsiPath -PathType Leaf)) {
            throw "MsiPath was not found: $MsiPath"
        }
    }
    'AIPackagingWinGet' {
        # -AppName is the PackageIdentifier; Description is enforced at the Param level.
    }
}
if ($missing.Count -gt 0) {
    throw "The '$Source' source requires the following parameter(s): $($missing -join ', ')."
}

# Preserve the original image input when it is a public URL. Resolve-ImagePath converts a URL into a
# downloaded LOCAL path (which Cloudpager needs), but Teams Adaptive Cards can only render a publicly
# reachable URL - so keep the original around for notifications.
$image_url_original = $null
if ($image_file_path -match '^(?i)https?://') { $image_url_original = $image_file_path }

# Resolve an image URL to a local path (applies to every mode; no-op for local paths / empty).
$image_file_path = Resolve-ImagePath $image_file_path

# ------------------------------------------------------------------------------------
# NIP-only prerequisites (CreateJson.ps1 / studio-nip.ps1 presence + self-update).
# These are irrelevant to the AI Packaging modes, so only run them for the NIP sources.
# ------------------------------------------------------------------------------------

if ($Source -in @('Evergreen','WinGet','Chocolatey')) {

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
}

# ------------------------------------------------------------------------------------
# Configuration - enter values for these variables before using the script
# ------------------------------------------------------------------------------------

$skey = "<Add SubscriptionKey>"

# Fail early with a clear message if the Subscription Key has not been set.
if ([string]::IsNullOrWhiteSpace($skey) -or $skey -match '^<.*>$') {
    throw "Set `$skey to your Cloudpager Subscription Key in the Configuration section before running this script."
}

Set-CloudpagerSubscriptionKey -Subscriptionkey $skey -Force

$AppsDashboardURL = '<Add Cloudpager Apps Page URL>'

# ----------------------------------------------------------------------------------
# Teams notifications (OPTIONAL - disabled by default).
# To enable: (1) paste your Teams "Workflows" (Power Automate) trigger URL into $TeamsURI below,
#            (2) uncomment the notification block at the very bottom of this script.
# The legacy Office 365 connector webhooks are being retired, so use a Workflows trigger URL.
# ----------------------------------------------------------------------------------
$TeamsURI = "<Add Teams Workflow URL>"

# Optional fallback banner image for the Teams card (public URL). Used when the run's -image_file_path
# was NOT supplied as a public URL. Leave blank to send the card without an image.
$TeamsImageURL = ''

# Remove the comments for the next 2 lines and supply your OpenAI API key to use the OpenAI API as part of the script.
#$gptkey = ConvertTo-SecureString "<OpenAI-API-Key>" -AsPlainText -Force
#Set-OpenAIKey -key $gptKey

# ------------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------------

# Publish tracking - each source sets these on a successful publish so a single Teams
# notification can be sent after the switch, regardless of which mode ran.
$Published      = $false
$PublishName    = $null
$PublishVersion = $null

switch ($Source) {

# ==================================================================================
# WinGet (NIP packaging)
# ==================================================================================
'WinGet' {

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

$Published = $true; $PublishName = $Name; $PublishVersion = $Version

If($WorkpodID){
Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$Name" -PublishComment "Added $Name $Version" -Confirm -Force
}
}
else
{
Write-Output "$Name is already published in your Cloudpager tenant."
}

} # end WinGet

# ==================================================================================
# Chocolatey (NIP packaging)
# ==================================================================================
'Chocolatey' {

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

If($Version -ne $Curversion -or $Curversion -eq $null){

New-Item "$NIPDirectory\Auto\Install.cmd"

Set-Content "$NIPDirectory\Auto\Install.cmd" "choco install $AppName -y"

.\CreateJson.ps1 -Filepath "$NIPDirectory\Auto\Install.cmd" -Description $Description -Name $Name -Arguments " " -StudioCommandLine $CommandLine -outputfolder "$NIPDirectory\Auto"

$config_file_path = Get-ChildItem -Path "$NIPDirectory\Auto" -Filter *.json | ForEach-Object{$_.FullName}

.\studio-nip.ps1 -config_file_path $config_file_path

$PackageFile = Get-ChildItem -Path "$NIPDirectory\Auto" -Filter *.stp | ForEach-Object{$_.FullName}

Add-CloudpagerApplication -SubscriptionKey $skey -Filepath $PackageFile -Name $Name -AppVersion $Version -Publisher $publisher -ImagePath $image_file_path -Description $Description -PublishComment "Uploaded using API" -Force

$Published = $true; $PublishName = $Name; $PublishVersion = $Version
}

} # end Chocolatey

# ==================================================================================
# AI Packaging - WinGet catalog (server-side packaging, no NIP)
# ==================================================================================
'AIPackagingWinGet' {

Try
{
    winget search "$AppName" | Out-Null
}
Catch
{
    Write-Error "An error has occurred retrieving data for $AppName from the Windows Package Manager. Ensure the App Installer app is installed, if errors continue try to run a query manually using winget search $AppName"
}

# -AppName is the WinGet PackageIdentifier. Derive a display Name and version from winget.
$pkgInfo = Get-WinGetPackageInfo -Identifier $AppName

if ($null -eq $pkgInfo) {
    Write-Output "Could not resolve '$AppName' from winget; falling back to the identifier as the display name."
    $Name = $AppName
}
else {
    $Name = $pkgInfo.Name
    Write-Output "$AppName found in WinGet! Name: $($pkgInfo.Name), Id: $($pkgInfo.Id), Version: $($pkgInfo.Version)"
}

# Version to publish / compare against the tenant: explicit override wins, else winget value.
if ($AppVersion) { $Version = $AppVersion } elseif ($pkgInfo) { $Version = $pkgInfo.Version }

Try
{
    Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion
}
Catch
{
    Write-Error "An error has occurred retrieving data for $Name from Cloudpager. Ensure the Cloudpager API is installed."
}

$Curversion = Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion
$Curversion = $Curversion | measure -Maximum | select -ExpandProperty Maximum

If($Version -ne $Curversion -or $Curversion -eq $null){

    Write-Output "AI Packaging (WinGet) '$Name' $Version ..."

    $catalogEntry = Get-CloudpagerWinGetCatalog -PackageIdentifier $AppName

    # The WinGet catalog object already contains the PackageIdentifier, application name,
    # publisher and package version required by the Cloudpager AI Packaging API.
    # Do NOT override Name/AppVersion here: newer Cloudpager API versions can reject
    # catalog-backed creates with HTTP 400 when those catalog-owned fields are supplied.
    if ($null -eq $catalogEntry) {
        throw "Cloudpager WinGet catalog did not return an entry for '$AppName'."
    }

    # Protect against an unexpectedly broad result. The PackageIdentifier query should
    # resolve to one catalog entry before it is piped into Add-CloudpagerApplication.
    $catalogEntries = @($catalogEntry)
    if ($catalogEntries.Count -ne 1) {
        throw "Expected exactly one Cloudpager WinGet catalog entry for '$AppName' but received $($catalogEntries.Count)."
    }
    $catalogEntry = $catalogEntries[0]

    $addParams = @{
        SubscriptionKey = $skey
        Description     = $Description
        PublishComment  = "Uploaded using API (AI Packaging - WinGet)"
        Force           = $true
    }

    # ImagePath is application metadata rather than WinGet catalog metadata, so retain it
    # when the caller supplied an icon. Publisher is intentionally NOT overridden here.
    if ($image_file_path) { $addParams.ImagePath = $image_file_path }

    try {
        $catalogEntry | Add-CloudpagerApplication @addParams -ErrorAction Stop
    }
    catch {
        $message = $_.Exception.Message
        throw "Failed to create Cloudpager AI-packaged WinGet application '$AppName'. Cloudpager returned: $message"
    }

    $Published = $true; $PublishName = $Name; $PublishVersion = $Version

    Write-Output "$Name $Version has been published in Cloudpager!"

    If($WorkpodID){
    Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$Name" -PublishComment "Added $Name $Version" -Confirm -Force
    }
}
else
{
Write-Output "$Name is already published in your Cloudpager tenant."
}

} # end AIPackagingWinGet

# ==================================================================================
# AI Packaging - supplied MSI (server-side packaging, no NIP)
# ==================================================================================
'AIPackagingMsi' {

$Name = $AppName

# Version: explicit override wins, else auto-extract ProductVersion from the MSI.
if ($AppVersion) {
    $Version = $AppVersion
}
else {
    Try
    {
        $Version = Get-MsiProperty -Path $MsiPath -Property 'ProductVersion'
    }
    Catch
    {
        Write-Error "Failed to read ProductVersion from the MSI at $MsiPath. Supply -AppVersion to continue. $_"
    }
}

if (-not $Version) {
    throw "Could not determine a version for the MSI. Supply -AppVersion explicitly."
}

# Publisher fallback: use the MSI Manufacturer property if the user did not supply one.
if (-not $Publisher) {
    Try { $Publisher = Get-MsiProperty -Path $MsiPath -Property 'Manufacturer' } Catch { }
}

Try
{
    Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion
}
Catch
{
    Write-Error "An error has occurred retrieving data for $Name from Cloudpager. Ensure the Cloudpager API is installed."
}

$Curversion = Get-CloudpagerApplication -SubscriptionKey $skey | Where-Object{$_.Name -like $Name} | Select -ExpandProperty AppVersion
$Curversion = $Curversion | measure -Maximum | select -ExpandProperty Maximum

If($Version -ne $Curversion -or $Curversion -eq $null){

    Write-Output "AI Packaging (MSI) '$Name' $Version from $MsiPath ..."

    $addParams = @{
        SubscriptionKey = $skey
        Filepath        = $MsiPath
        Name            = $Name
        AppVersion      = $Version
        Description     = $Description
        PublishComment  = "Uploaded using API (AI Packaging - MSI)"
        Force           = $true
    }
    if ($Publisher)       { $addParams.Publisher         = $Publisher }
    if ($Arguments)       { $addParams.InstallerArguments = $Arguments }
    if ($image_file_path) { $addParams.ImagePath         = $image_file_path }

    Add-CloudpagerApplication @addParams

    $Published = $true; $PublishName = $Name; $PublishVersion = $Version

    Write-Output "$Name $Version is now available in Cloudpager!"

    If($WorkpodID){
    Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$Name" -PublishComment "Added $Name $Version" -Confirm -Force
    }
}
else
{
Write-Output "$Name is already published in your Cloudpager tenant."
}

} # end AIPackagingMsi

# ==================================================================================
# Evergreen (NIP packaging) - default source
# ==================================================================================
'Evergreen' {

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
#$Publisher = ai "What vendor makes $FriendlyName? Just return the short name, no other text and no period." | Out-String

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
$Published = $true; $PublishName = $FriendlyName; $PublishVersion = $LatestVersion
If($WorkpodID){
Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$FriendlyName" -PublishComment "Added $FriendlyName $LatestVersion" -Confirm -Force
}

}
else
{

#Remove comment for the line below and change Description parameter to let OpenAI API populate the Publisher for you.
#$Description = ai "What is $Publisher $AppName in 30 words or less." | Out-String

.\CreateJson.ps1 -Filepath $DownloadFilePath -Description $Description -Name $FriendlyName -Arguments $Arguments -RegistryExclusions $registryexclusions -FileExclusions $fileexclusion -StudioCommandLine $CommandLine -outputfolder "$ProjectFolder\Output"

$config_file_path = Get-ChildItem -Path "C:\NIP_Software\Auto" -Filter *.json | ForEach-Object{$_.FullName}

.\studio-nip.ps1 -config_file_path $config_file_path

$PackageFile = Get-ChildItem -Path "$ProjectFolder\Output" -Filter *.stp | ForEach-Object{$_.FullName}

Add-CloudpagerApplication -SubscriptionKey $skey -Filepath $PackageFile -Name $FriendlyName -AppVersion $LatestVersion -Publisher $publisher -ImagePath $image_file_path -Description $Description -PublishComment "Uploaded using API" -Force

$Published = $true; $PublishName = $FriendlyName; $PublishVersion = $LatestVersion

Write-Output "$AppName $LatestVersion has been published in Cloudpager!"

If($WorkpodID -and $WinGet -ne $True){
Set-CloudpagerWorkpod -Subscriptionkey $skey -WorkpodID $WorkpodID -Applications "$FriendlyName" -PublishComment "Added $FriendlyName $LatestVersion" -Confirm -Force
}
}
}
else
{
Write-Output "Latest version of $AppName is already published in Cloudpager"
}

} # end Evergreen

} # end switch

# ------------------------------------------------------------------------------------
# Teams notification (OPTIONAL - DISABLED BY DEFAULT).
# Fires for ANY source after a successful publish. To enable notifications:
#   1. Set $TeamsURI in the Configuration section to your Teams Workflows trigger URL.
#   2. Uncomment the block below (select the lines and remove the leading '#').
# It is guarded so it stays silent unless something was actually published AND $TeamsURI
# is set to a real value (not the "<...>" placeholder).
# ------------------------------------------------------------------------------------
#if ($Published -and $TeamsURI -and $TeamsURI -notmatch '^<.*>$') {
#
#    # Prefer the run's public image URL; fall back to the configured banner image.
#    $notifyImage = if ($image_url_original) { $image_url_original } elseif ($TeamsImageURL) { $TeamsImageURL } else { $null }
#
#    $teamsParams = @{
#        WebhookUrl = $TeamsURI
#        Title      = 'Application Update'
#        Message    = "$PublishName $PublishVersion has been published in Cloudpager."
#        LinkUrl    = $AppsDashboardURL
#        ButtonText = 'View in Cloudpager'
#    }
#    # Only add ImageUrl when a public URL exists (a local path cannot be rendered by Teams,
#    # and the parameter is [ValidateNotNullOrEmpty] so it must not be passed as $null).
#    if ($notifyImage) { $teamsParams.ImageUrl = $notifyImage }
#
#    try {
#        Send-TeamsAdaptiveCard @teamsParams -Verbose
#    }
#    catch {
#        Write-Warning $_.Exception.Message
#    }
#}
