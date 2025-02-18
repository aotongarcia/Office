function Get-ConfigurationManagerSiteCode {
    <#
    .SYNOPSIS
        Retrieves the site code from a specified Configuration Manager site server.

    .DESCRIPTION
        This function queries the SMS Provider on a specified Configuration Manager site server
        to determine the site code. It uses CIM to connect to the root\SMS namespace and
        retrieves the site code from the SMS_ProviderLocation class.

    .PARAMETER SiteServer
        The fully qualified domain name (FQDN) or hostname of the Configuration Manager site server
        that hosts the SMS Provider.

    .OUTPUTS
        System.String
        Returns the site code as a string value.

    .EXAMPLE
        PS C:\> Get-ConfigurationManagerSiteCode -SiteServer "CM01.contoso.com"
        P01
        This example retrieves the site code from the Configuration Manager server CM01.contoso.com.
    #>

    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SiteServer
    )

    begin {
        Write-Verbose -Message "Determining Site Code for Site Server: $siteServer"
    }

    process {
        try {
            $siteCode = Get-CimInstance -Namespace 'root\SMS' -Class SMS_ProviderLocation -ComputerName $SiteServer -ErrorAction Stop -Verbose:$false | Where-Object { $_.ProviderForLocalSite -eq $true } | Select-Object -First 1 -ExpandProperty SiteCode

            if ($null -eq $siteCode) {
                throw "No local site provider found on server '$SiteServer'"
            }
        }
        catch {
            Write-Warning -Message 'Unable to determine site code from specified Configuration Manager site server. Ensure the SMS Provider is installed.'
            throw $($_.Exception.Message)
        }
    }

    end {
        $siteCode
    }
}

function Import-ConfigurationManagerModule {
    <#
    .SYNOPSIS
        Imports the Configuration Manager PowerShell module and sets up the required PSDrive.

    .DESCRIPTION
        This function attempts to import the Configuration Manager PowerShell module using two methods:
        1. Direct import using Import-Module
        2. Alternative import using the SMS_ADMIN_UI_PATH environment variable

        After successful import, it creates a PSDrive for the specified site code if it doesn't exist.

    .PARAMETER SiteCode
        The Configuration Manager site code (e.g., 'P01' or 'PS1') that will be used
        for the PSDrive creation.

    .PARAMETER SiteServer
        The fully qualified domain name (FQDN) or hostname of the Configuration Manager
        site server that hosts the SMS Provider.

    .EXAMPLE
        PS C:\> Import-ConfigurationManagerModule -SiteCode "PS1" -SiteServer "SCCM01.contoso.com"
        Imports the ConfigMgr module and creates a PS1: drive mapped to SCCM01.contoso.com
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SiteCode,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SiteServer
    )

    begin {}

    process {
        try {
            Import-Module -Name ConfigurationManager -ErrorAction Stop -Verbose:$false
            Write-Verbose 'Successfully imported ConfigurationManager module'
        }
        catch {
            Write-Warning 'Direct import failed, attempting alternative import method'
            try {
                if (-not $env:SMS_ADMIN_UI_PATH) {
                    throw 'SMS_ADMIN_UI_PATH environment variable not found'
                }

                $configManagerModulePath = Join-Path -Path (($env:SMS_ADMIN_UI_PATH).Substring(0, $env:SMS_ADMIN_UI_PATH.Length - 5)) -ChildPath 'ConfigurationManager.psd1'

                if (-not (Test-Path -Path $configManagerModulePath)) {
                    throw "ConfigurationManager module not found at: $configManagerModulePath"
                }

                Write-Verbose "Importing ConfigurationManager module from: $configManagerModulePath"
                Import-Module $configManagerModulePath -Force -ErrorAction Stop -Verbose:$false

                # Create PSDrive if it doesn't exist
                if ($null -eq (Get-PSDrive -Name $siteCode -ErrorAction SilentlyContinue)) {
                    Write-Verbose "Creating PSDrive for site code: $siteCode"
                    New-PSDrive -Name $siteCode -PSProvider CMSite -Root $siteServer -ErrorAction Stop | Out-Null
                }

                Write-Verbose 'Successfully imported ConfigurationManager module using alternative method'
            }
            catch {
                Write-Warning 'Failed to load the ConfigurationManager module'
                throw $_.Exception.Message
            }
        }
    }

    end {}
}

function Get-OfficeDeploymentToolDownloadUrl {
    <#
    .SYNOPSIS
        Retrieves the download URL for the latest Office Deployment Tool.

    .DESCRIPTION
        This function scrapes the Microsoft Download Center page for the Office Deployment Tool
        to obtain the latest download URL. It parses the HTML content to find the download link
        associated with the ODT executable.

    .OUTPUTS
        [PSCustomObject] Returns an object containing:
        - DownloadUrl: The direct download URL for the Office Deployment Tool
        - FileName: The name of the executable file

    .EXAMPLE
        PS C:\> $odtInfo = Get-OfficeDeploymentToolDownloadUrl
        PS C:\> $odtInfo.DownloadUrl
        https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_12345-20000.exe

    .NOTES
        Known Issues:
        - The parsing logic may need updates if Microsoft changes their download page structure
        - The download URL (id=49117) may change in future Microsoft website updates
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param()

    begin {
        # This URL can change in the future
        $downloadPageUrl = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
    }

    process {
        try {
            Write-Verbose "Parsing Office Deployment Tool download page: $downloadPageUrl"
            $webResponse = Invoke-WebRequest -Uri $downloadPageUrl -UseBasicParsing -ErrorAction Stop -Verbose:$false

            # This parsing might need to be updated if the page changes. Currently, the download link is found by searching for the 'Download' span element and extracting the href attribute
            $downloadUrl = ($webResponse.Links | Where-Object { $_.outerHTML -match '<span[^>]*>Download</span>' }).href


            if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
                throw 'Parsing of download page failed, unable to find download link. Please inspect the page structure and update the script accordingly.'
            }
        }
        catch {
            Write-Warning 'Failed to parse Office Deployment Tool download page'
            throw $($_.Exception.Message)
        }
    }
    end {
        # Return download information
        [PSCustomObject]@{
            DownloadUrl = $downloadUrl
            FileName    = Split-Path -Path $downloadUrl -Leaf
        }
    }
}

function Start-OfficeDeploymentToolDownload {
    <#
    .SYNOPSIS
        Downloads the Office Deployment Tool executable from Microsoft.

    .DESCRIPTION
        This function downloads the Office Deployment Tool (ODT) executable from Microsoft's servers.
        It creates the destination directory if needed, downloads the file, and verifies the download
        was successful by checking file existence and size.

    .PARAMETER OfficeDeploymentToolDownloadInformation
        A PSCustomObject containing:
        - DownloadUrl: The direct download URL for the Office Deployment Tool
        - FileName: The name of the executable file to be downloaded

    .PARAMETER DestinationPath
        The full file system path where the Office Deployment Tool executable will be downloaded.
        If the directory doesn't exist, it will be created.

    .PARAMETER TimeoutSeconds
        Optional. The number of seconds to wait for the download to complete before timing out.
        Default value is 300 seconds (5 minutes).

    .OUTPUTS
        System.String
        Returns the full path to the downloaded Office Deployment Tool executable.

    .EXAMPLE
        PS C:\> $downloadInfo = @{
            DownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_12345-20000.exe"
            FileName = "officedeploymenttool_12345-20000.exe"
        }
        PS C:\> Start-OfficeDeploymentToolDownload -OfficeDeploymentToolDownloadInformation $downloadInfo -DestinationPath "C:\Temp\ODT"
        Downloads the ODT executable to C:\Temp\ODT\officedeploymenttool_12345-20000.exe
    #>

    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject]
        $OfficeDeploymentToolDownloadInformation,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DestinationPath,

        [Parameter(Mandatory = $false)]
        [int]
        $TimeoutSeconds = 300
    )

    begin {
        Write-Host 'Starting Office Deployment Tool self-extracting executable download' -ForegroundColor Cyan
        $downloadFilePath = Join-Path -Path $destinationPath -ChildPath $OfficeDeploymentToolDownloadInformation.FileName
    }

    process {
        try {
            # Ensure destination directory exists
            if (-not(Test-Path -Path $destinationPath)) {
                Write-Verbose "Creating directory: $destinationPath"
                New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
            }

            # Configure download Parameters
            $downloadParams = @{
                Uri             = $officeDeploymentToolDownloadInformation.DownloadUrl
                OutFile         = $downloadFilePath
                Method          = 'GET'
                UseBasicParsing = $true
                TimeoutSec      = $TimeoutSeconds
                ErrorAction     = 'Stop'
                Verbose         = $false
            }

            Write-Verbose "Downloading Office Deployment Tool self-extracting executable from: $($officeDeploymentToolDownloadInformation.DownloadUrl)"
            Write-Verbose "Destination: $downloadFilePath"

            # Start download with progress
            Invoke-RestMethod @downloadParams

            # Verify download
            if (Test-Path -Path $downloadFilePath) {
                $downloadedFile = Get-Item -Path $downloadFilePath

                if ($downloadedFile.Length -gt 0) {
                    Write-Verbose "Successfully downloaded Office Deployment Tool self-extracting executable (Size: $([math]::Round($downloadedFile.Length/1MB, 2)) MB)"
                }
                else {
                    throw 'Downloaded failed - the file is empty'
                }
            }
            else {
                throw 'Download failed - file not found'
            }
        }
        catch {
            Write-Warning 'Failed to download Office Deployment Tool'
            throw $($_.Exception.Message)
        }
    }

    end {
        $downloadFilePath
    }
}

function Start-OfficeDeploymentToolExtraction {
    <#
    .SYNOPSIS
        Extracts the Office Deployment Tool to a version-specific directory.

    .DESCRIPTION
        This function extracts the Office Deployment Tool (ODT) executable to a versioned directory.
        It performs the following tasks:
        - Creates a version-specific extraction directory
        - Extracts the ODT files using the /quiet /extract switches
        - Verifies the extraction by checking for setup.exe
        - Returns extraction details including paths and version information

    .PARAMETER OfficeDeploymentToolExecutable
        The full path to the downloaded Office Deployment Tool executable file.
        This should be the self-extracting exe downloaded from Microsoft.

    .PARAMETER ExtractionPath
        The base path where the Office Deployment Tool will be extracted.
        A version-specific subdirectory will be created under this path.

    .OUTPUTS
        [PSCustomObject] containing:
        - ExtractionPath: The full path to the version-specific extraction directory
        - SetupFilePath: The full path to the extracted setup.exe
        - LatestOfficeDeploymentToolVersion: The version number of the extracted ODT

    .EXAMPLE
        PS C:\> $extractionInfo = Start-OfficeDeploymentToolExtraction `
            -OfficeDeploymentToolExecutable "C:\Temp\officedeploymenttool.exe" `
            -ExtractionPath "C:\ODT"

        Extracts the ODT to a version-specific folder under C:\ODT and returns the extraction details.
    #>

    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $OfficeDeploymentToolExecutable,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ExtractionPath
    )

    begin {
        # Construct a versioned extraction path using the version of the Office Deployment Tool executable
        $officeDeploymentToolVersionedExtractionPath = Join-Path -Path $extractionPath -ChildPath (Get-Item $officeDeploymentToolExecutable).VersionInfo.ProductVersion
        # Construct the "setup.exe" path
        $setupFilePath = Join-Path -Path $officeDeploymentToolVersionedExtractionPath -ChildPath 'setup.exe'
        # Construct extraction arguments
        $extractionArguments = "/quiet /extract:`"$officeDeploymentToolVersionedExtractionPath`""
    }

    process {
        try {
            if (-not(Test-Path -Path $officeDeploymentToolVersionedExtractionPath)) {
                Write-Verbose "Creating directory: $officeDeploymentToolVersionedExtractionPath"
                New-Item -Path $officeDeploymentToolVersionedExtractionPath -ItemType Directory -Force | Out-Null
            }

            # Extract Office Deployment Tool files
            Write-Verbose "Extracting Office Deployment Tool to: $officeDeploymentToolVersionedExtractionPath"
            Start-Process -FilePath $officeDeploymentToolExecutable -ArgumentList $extractionArguments -Wait -ErrorAction Stop
            # Verify extraction
            if (Test-Path -Path $setupFilePath) {
                Write-Verbose "Successfully extracted Office Deployment Tool to: $officeDeploymentToolVersionedExtractionPath"

            }
            else {
                throw 'Extraction failed - "setup.exe" not found'
            }
        }
        catch {
            Write-Warning 'Failed to extract Office Deployment Tool'
            throw $($_.Exception.Message)
        }
    }

    end {
        [PSCustomObject] @{
            ExtractionPath                    = $officeDeploymentToolVersionedExtractionPath
            SetupFilePath                     = $setupFilePath
            LatestOfficeDeploymentToolVersion = [version](Get-Item -Path "$officeDeploymentToolVersionedExtractionPath\setup.exe").VersionInfo.ProductVersion
        }
    }
}

Write-Host 'Initiating Office application content update process' -ForegroundColor Cyan

$officeDeploymentToolDownloadUrl = Get-OfficeDeploymentToolDownloadUrl

$officeDeploymentToolExecutableDownload = Start-OfficeDeploymentToolDownload -OfficeDeploymentToolDownloadInformation $officeDeploymentToolDownloadUrl -DestinationPath $officeDeploymentToolExtractionPath

$officeDeploymentToolInformation = Start-OfficeDeploymentToolExtraction -OfficeDeploymentToolExecutable $officeDeploymentToolExecutableDownload -ExtractionPath $officeDeploymentToolExtractionPath

# Compare the current version of the Office Deployment Tool executable with the latest version and update if necessary
try {
    # Get the current version of the Office Deployment Tool executable
    [version]$currentOfficeDeploymentToolVersion = (Get-Item -Path (Join-Path $officeContentPath -ChildPath 'setup.exe') -ErrorAction Stop).VersionInfo.ProductVersion

    if ($officeDeploymentToolInformation.LatestOfficeDeploymentToolVersion -gt $currentOfficeDeploymentToolVersion) {
        Write-Warning 'Newer version of Office Deployment Tool available - updating content source'

        Write-Verbose "Current Office Deployment Tool version: $currentOfficeDeploymentToolVersion"
        Write-Verbose "Latest Office Deployment Tool version: $($officeDeploymentToolInformation.LatestOfficeDeploymentToolVersion)"

        try {
            Copy-Item -Path $officeDeploymentToolInformation.SetupFilePath -Destination $officeContentPath -Force -ErrorAction Stop
        }
        catch {
            Write-Warning 'There was an error updating the Office Deployment Tool executable in application content source'
            throw $($_.Exception.Message)
        }

        Write-Host 'Successfully updated Office Deployment Tool executable in application content source' -ForegroundColor Cyan
    }
    else {
        Write-Host 'Office Deployment Tool version is up to date' -ForegroundColor Green
    }
}
catch {
    Write-Warning 'Failed to compare Office Deployment Tool versions'
    throw $($_.Exception.Message)
}

# Clean up extracted files
try {
    Write-Verbose "Removing extracted Office Deployment Tool files from: $($officeDeploymentToolInformation.ExtractionPath)"

    Remove-Item -Path $officeDeploymentToolInformation.ExtractionPath -Recurse -Force -ErrorAction Stop
    Remove-Item -Path $officeDeploymentToolExecutableDownload -Force -ErrorAction Stop
}
catch {
    Write-Warning "Failed to remove extracted Office Deployment Tool files. Please remove them manually - $($_.Exception.Message)"
}

# Detect current version information for existing Office application content
try {
    Write-Verbose 'Detecting current version information for existing Office application'

    $officeDataFolderRoot = Join-Path -Path $officeContentPath -ChildPath 'office\data'
    $currentOfficeDataFolder = Get-ChildItem -Path $officeDataFolderRoot -Directory -ErrorAction Stop
    $currentOfficeDataFile = Get-ChildItem -Path $officeDataFolderRoot -Filter 'v*_*.cab' -ErrorAction Stop
}
catch {
    Write-Warning 'Failed to detect current version information for existing Office application'
    throw $($_.Exception.Message)
}

# Update Office application content based on configuration file
try {
    Write-Verbose "Updating application content based on configuration file: $officeConfigurationFile"

    $setupExecutablePath = Join-Path -Path $officeContentPath -ChildPath 'setup.exe'

    Start-Process -FilePath $setupExecutablePath -ArgumentList "/download $officeConfigurationFile" -WorkingDirectory $officeContentPath -Wait -NoNewWindow -ErrorAction Stop
}
catch {
    Write-Warning 'Failed to update the Office application content.'
    throw $($_.Exception.Message)
}

# Check if new content was downloaded and remove old content if necessary
try {
    Write-Verbose 'Checking if new content was downloaded and removing old content if necessary.'

    # Check if more than one data folder exists
    if ((Get-ChildItem -Path $officeDataFolderRoot -Directory).Count -ge 2) {
        Write-Verbose "Removing old content from: $($officeDataFolderRoot.Name)"
        # Remove old data folder
        Write-Verbose "Removing old data folder: $($currentOfficeDataFolder.Name)"
        Remove-Item -Path $currentOfficeDataFolder.FullName -Recurse -Force -ErrorAction Stop
        # Remove old cab file
        Write-Verbose "Removing old cab file: $($currentOfficeDataFile.Name)"
        Remove-Item -Path $currentOfficeDataFile.FullName -Force -ErrorAction Stop

        $latestOfficeDataFolder = Get-ChildItem -Path $officeDataFolderRoot -Directory -ErrorAction Stop

        Write-Host "Successfully updated Office application content from $($currentOfficeDataFolder.Name) to $($latestOfficeDataFolder.Name)" -ForegroundColor Green
    }

    # If only one data folder exists, set the latest data folder to the current data folder
    else {
        $latestOfficeDataFolder = $currentOfficeDataFolder
    }
}
catch {
    Write-Warning 'Failed to remove old content from the Office application content source.'
    throw $($_.Exception.Message)
}

$siteCode = Get-ConfigurationManagerSiteCode -SiteServer $siteServer

Import-ConfigurationManagerModule -SiteServer $siteServer -SiteCode $siteCode

# Set the current location to the Configuration Manager drive
try {
    Set-Location -Path ($siteCode + ':') -ErrorAction Stop -Verbose:$false
}
catch {
    Write-Warning 'Failed to set the current location to the Configuration Manager drive'
    throw $($_.Exception.Message)
}

# Get the application and application deployment type details
try {
    $officeApplication = Get-CMApplication -Name $officeApplicationName -Fast -ErrorAction Stop -Verbose:$false

    if ($null -eq $officeApplication) {
        throw "No application found with the name: `"$officeApplicationName`". Please make sure `"$officeApplicationName`" is a valid application name in Configuration Manager."
    }

    $officeApplicationDeploymentType = Get-CMDeploymentType -ApplicationName $officeApplicationName -ErrorAction Stop -Verbose:$false

    if ($null -eq $officeApplicationDeploymentType) {
        throw "No deployment type found for the Office application: `"$officeApplicationName`". Please make sure `"$officeApplicationName`" is a valid Office application name in Configuration Manager or that it has a deployment type."
    }
}
catch {
    Write-Warning 'Failed to retrieve the Office application information'
    throw $($_.Exception.Message)
}

# Update application metadata and detection method
if ($PSBoundParameters.ContainsKey('UpdateConfigurationManagerDetectionMethod')) {
    Write-Host "Updating application metadata and detection method for: `"$officeApplicationName`"" -ForegroundColor Magenta

    # Update the software version for the Office application metadata
    if ($officeApplication.SoftwareVersion -ne $latestOfficeDataFolder.Name) {
        try {
            Write-Verbose 'Updating Office application metadata'
            Set-CMApplication -InputObject $officeApplication -SoftwareVersion $($latestOfficeDataFolder.Name) -ErrorAction Stop -Verbose:$false
        }
        catch {
            Write-Warning 'Failed to update the software version for the Office application metadata'
            throw $($_.Exception.Message)
        }
    }

    # Create a new registry detection clause for the Office application
    try {
        Write-Verbose "Creating new registry detection clause deployment type for: `"$($officeApplicationDeploymentType.LocalizedDisplayName)`""

        $detectionClauseArguments = @{
            ExpressionOperator = 'GreaterEquals'
            Hive               = 'LocalMachine'
            KeyName            = 'Software\Microsoft\Office\ClickToRun\Configuration'
            PropertyType       = 'Version'
            ValueName          = 'VersionToReport'
            ExpectedValue      = $latestOfficeDataFolder.Name
            Value              = $true
            ErrorAction        = 'Stop'
            Verbose            = $false
        }

        $newRegistryDetectionClause = New-CMDetectionClauseRegistryKeyValue @detectionClauseArguments
    }
    catch {
        Write-Warning 'Failed to create a new registry detection clause'
        throw $($_.Exception.Message)
    }

    # Update the detection method for the Office application deployment type
    try {
        Write-Verbose "Updating detection method for deployment type for: `"$($officeApplicationDeploymentType.LocalizedDisplayName)`""

        Write-Verbose "Retrieving current registry detection clause from the deployment type: `"$($officeApplicationDeploymentType.LocalizedDisplayName)`""

        $currentRegistryDetectionClause = ([xml]$officeApplicationDeploymentType.SDMPackageXML).AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.SimpleSetting | Where-Object { $_.DataType -eq 'Version' } | Select-Object -ExpandProperty LogicalName

        Write-Verbose "Current registry detection clause name: $currentRegistryDetectionClause"

        Write-Verbose 'Removing current registry detection clause and adding new detection clause'

        switch ($officeApplicationDeploymentType.Technology) {
            'MSI' {
                Set-CMMsiDeploymentType -InputObject $officeApplicationDeploymentType -RemoveDetectionClause $currentRegistryDetectionClause -AddDetectionClause $newRegistryDetectionClause -ErrorAction Stop -Verbose:$false
            }
            'Script' {
                Set-CMScriptDeploymentType -InputObject $officeApplicationDeploymentType -RemoveDetectionClause $currentRegistryDetectionClause -AddDetectionClause $newRegistryDetectionClause -ErrorAction Stop -Verbose:$false
            }
            default {
                throw "Unsupported deployment type technology: $($officeApplicationDeploymentType.Technology)"
            }
        }
    }
    catch {
        Write-Warning "Failed to update the detection method for deployment type: $($officeApplicationDeploymentType.LocalizedDisplayName)"
        throw $($_.Exception.Message)
    }
}

try {
    Write-Host "Beginning content distribution for Office application: '$officeApplicationName'" -ForegroundColor Magenta

    Update-CMDistributionPoint -ApplicationName $officeApplicationName -DeploymentTypeName $officeApplicationDeploymentType.LocalizedDisplayName -ErrorAction Stop -Verbose:$false

    Write-Host "Successfully started the content distribution for Office application: `"$officeApplicationName`"" -ForegroundColor Green
}
catch {
    Write-Warning 'Failed to start the content distribution for the Office application'
    throw $($_.Exception.Message)
}

Write-Host 'Office application content update process completed' -ForegroundColor Cyan

Set-Location $env:SystemDrive
