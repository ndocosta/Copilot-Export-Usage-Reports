<#
.SYNOPSIS
    Exports Microsoft 365 Copilot usage reports from Graph API to CSV and uploads to SharePoint Online.

.DESCRIPTION
    This script authenticates to Microsoft Graph API using PnP PowerShell with certificate-based authentication,
    retrieves Copilot usage reports, exports them to CSV files locally, and uploads them to SharePoint Online.

.PARAMETER ConfigFile
    Path to the configuration JSON file. Defaults to config.json in the script directory.

.EXAMPLE
    .\Export-CopilotUsageReports.ps1
    .\Export-CopilotUsageReports.ps1 -ConfigFile "C:\Config\myconfig.json"

.NOTES
    Author: Microsoft
    Date: January 5, 2026
    Requires: PnP.PowerShell module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigFile = (Join-Path $PSScriptRoot "config.m365cpi15450148.json")
)

#region Functions

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Error' { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
    
    # Append to log file (only if config and LocalExportPath are available)
    if ($script:config -and $script:config.LocalExportPath -and (Test-Path $script:config.LocalExportPath -PathType Container)) {
        try {
            $logFile = Join-Path $script:config.LocalExportPath "CopilotExport_$(Get-Date -Format 'yyyyMMdd').log"
            Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
        }
        catch {
            # Silently fail if we can't write to log file
        }
    }
}

function Get-CopilotUsageReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReportType,
        
        [Parameter(Mandatory = $false)]
        [int]$Period = 7
    )
    
    try {
        Write-Log "Retrieving $ReportType report for the last $Period days..."
        
        # Old Endpoints commented out. Now we should use /copilot/ namespace
        # # Graph API endpoints for Copilot usage reports
        # $endpoint = switch ($ReportType) {
        #     'UserDetail' { "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D$Period')" }
        #     'UserCountsSummary' { "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountSummary(period='D$Period')" }
        #     'UserCountsTrend' { "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountTrend(period='D$Period')" }
        #      default { throw "Unknown report type: $ReportType" }
        # }

        
        # Graph API endpoints for Copilot usage reports
        # TODO: update to v1.0 when available
        # Documentation: https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail?pivots=graph-preview
        $endpoint = switch ($ReportType) {
            'UserDetail' { "https://graph.microsoft.com/beta/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D$Period')" }
            'UserCountsSummary' { "https://graph.microsoft.com/beta/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='D$Period')" }
            'UserCountsTrend' { "https://graph.microsoft.com/beta/copilot/reports/getMicrosoft365CopilotUserCountTrend(period='D$Period')" }
            default { throw "Unknown report type: $ReportType" }
        }
        
        # Make Graph API call using Invoke-PnPGraphMethod
        $response = Invoke-PnPGraphMethod -Url $endpoint -Method Get -ConsistencyLevelEventual
        
        if ($response.value) {
            Write-Log "Successfully retrieved $($response.value.Count) records for $ReportType" -Level Success
            return $response.value
        }
        else {
            Write-Log "No data returned for $ReportType" -Level Warning
            return @()
        }
    }
    catch {
        Write-Log "Error retrieving $ReportType report: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Export-ToCSV {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Data,
        
        [Parameter(Mandatory = $true)]
        [string]$ReportType,
        
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )
    
    try {
        if ($Data.Count -eq 0) {
            Write-Log "No data to export for $ReportType" -Level Warning
            return $null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        # Ensure export directory exists
        if (-not (Test-Path $ExportPath)) {
            New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
            Write-Log "Created export directory: $ExportPath"
        }
        
        # Process data based on report type for optimal human readability
        switch ($ReportType) {
            'UserDetail' {
                # Flatten user detail records - expand nested period details
                $flattenedData = foreach ($record in $Data) {
                    $baseRecord = [ordered]@{
                        'ReportRefreshDate'       = $record.reportRefreshDate
                        'UserPrincipalName'       = $record.userPrincipalName
                        'DisplayName'             = $record.displayName
                        'LastActivityDate'        = $record.lastActivityDate
                        'CopilotChatLastActivity' = $record.copilotChatLastActivityDate
                        'TeamsLastActivity'       = $record.microsoftTeamsCopilotLastActivityDate
                        'WordLastActivity'        = $record.wordCopilotLastActivityDate
                        'ExcelLastActivity'       = $record.excelCopilotLastActivityDate
                        'PowerPointLastActivity'  = $record.powerPointCopilotLastActivityDate
                        'OutlookLastActivity'     = $record.outlookCopilotLastActivityDate
                        'OneNoteLastActivity'     = $record.oneNoteCopilotLastActivityDate
                        'LoopLastActivity'        = $record.loopCopilotLastActivityDate
                    }
                    
                    # Add period details if present
                    if ($record.copilotActivityUserDetailsByPeriod) {
                        $baseRecord['ReportPeriod'] = $record.copilotActivityUserDetailsByPeriod.reportPeriod
                    }
                    
                    [PSCustomObject]$baseRecord
                }
                
                $fileName = "CopilotUsage_${ReportType}_${timestamp}.csv"
                $filePath = Join-Path $ExportPath $fileName
                $flattenedData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            }
            
            'UserCountsSummary' {
                # Expand the nested adoption metrics into separate columns
                $flattenedData = foreach ($record in $Data) {
                    $adoption = $record.adoptionByProduct
                    
                    [PSCustomObject][ordered]@{
                        'ReportRefreshDate'       = $record.reportRefreshDate
                        'ReportPeriod'            = $adoption.reportPeriod
                        'TeamsEnabledUsers'       = $adoption.microsoftTeamsEnabledUsers
                        'TeamsActiveUsers'        = $adoption.microsoftTeamsActiveUsers
                        'WordEnabledUsers'        = $adoption.wordEnabledUsers
                        'WordActiveUsers'         = $adoption.wordActiveUsers
                        'PowerPointEnabledUsers'  = $adoption.powerPointEnabledUsers
                        'PowerPointActiveUsers'   = $adoption.powerPointActiveUsers
                        'OutlookEnabledUsers'     = $adoption.outlookEnabledUsers
                        'OutlookActiveUsers'      = $adoption.outlookActiveUsers
                        'ExcelEnabledUsers'       = $adoption.excelEnabledUsers
                        'ExcelActiveUsers'        = $adoption.excelActiveUsers
                        'OneNoteEnabledUsers'     = $adoption.oneNoteEnabledUsers
                        'OneNoteActiveUsers'      = $adoption.oneNoteActiveUsers
                        'LoopEnabledUsers'        = $adoption.loopEnabledUsers
                        'LoopActiveUsers'         = $adoption.loopActiveUsers
                        'AnyAppEnabledUsers'      = $adoption.anyAppEnabledUsers
                        'AnyAppActiveUsers'       = $adoption.anyAppActiveUsers
                        'CopilotChatEnabledUsers' = $adoption.copilotChatEnabledUsers
                        'CopilotChatActiveUsers'  = $adoption.copilotChatActiveUsers
                    }
                }
                
                $fileName = "CopilotUsage_${ReportType}_${timestamp}.csv"
                $filePath = Join-Path $ExportPath $fileName
                $flattenedData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            }
            
            'UserCountsTrend' {
                # Expand the trend data - each date becomes a separate row
                $flattenedData = foreach ($record in $Data) {
                    foreach ($dailyData in $record.adoptionByDate) {
                        [PSCustomObject][ordered]@{
                            'ReportRefreshDate'       = $record.reportRefreshDate
                            'ReportPeriod'            = $record.reportPeriod
                            'ReportDate'              = $dailyData.reportDate
                            'TeamsEnabledUsers'       = $dailyData.microsoftTeamsEnabledUsers
                            'TeamsActiveUsers'        = $dailyData.microsoftTeamsActiveUsers
                            'WordEnabledUsers'        = $dailyData.wordEnabledUsers
                            'WordActiveUsers'         = $dailyData.wordActiveUsers
                            'PowerPointEnabledUsers'  = $dailyData.powerPointEnabledUsers
                            'PowerPointActiveUsers'   = $dailyData.powerPointActiveUsers
                            'OutlookEnabledUsers'     = $dailyData.outlookEnabledUsers
                            'OutlookActiveUsers'      = $dailyData.outlookActiveUsers
                            'ExcelEnabledUsers'       = $dailyData.excelEnabledUsers
                            'ExcelActiveUsers'        = $dailyData.excelActiveUsers
                            'OneNoteEnabledUsers'     = $dailyData.oneNoteEnabledUsers
                            'OneNoteActiveUsers'      = $dailyData.oneNoteActiveUsers
                            'LoopEnabledUsers'        = $dailyData.loopEnabledUsers
                            'LoopActiveUsers'         = $dailyData.loopActiveUsers
                            'AnyAppEnabledUsers'      = $dailyData.anyAppEnabledUsers
                            'AnyAppActiveUsers'       = $dailyData.anyAppActiveUsers
                            'CopilotChatEnabledUsers' = $dailyData.copilotChatEnabledUsers
                            'CopilotChatActiveUsers'  = $dailyData.copilotChatActiveUsers
                        }
                    }
                }
                
                $fileName = "CopilotUsage_${ReportType}_${timestamp}.csv"
                $filePath = Join-Path $ExportPath $fileName
                $flattenedData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            }
            
            default {
                # Generic handling for any other report types - flatten nested structures
                $flattenedData = foreach ($record in $Data) {
                    $flatRecord = [PSCustomObject]@{}
                    
                    foreach ($property in $record.PSObject.Properties) {
                        $value = $property.Value
                        
                        if ($null -eq $value) {
                            $flatRecord | Add-Member -MemberType NoteProperty -Name $property.Name -Value ""
                        }
                        elseif ($value -is [Array] -or $value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
                            $stringValue = ($value | ForEach-Object { 
                                    if ($_ -is [PSCustomObject] -or $_ -is [Hashtable]) {
                                        $_ | ConvertTo-Json -Compress
                                    }
                                    else {
                                        $_
                                    }
                                }) -join '; '
                            $flatRecord | Add-Member -MemberType NoteProperty -Name $property.Name -Value $stringValue
                        }
                        elseif ($value -is [PSCustomObject] -or $value -is [Hashtable]) {
                            $flatRecord | Add-Member -MemberType NoteProperty -Name $property.Name -Value ($value | ConvertTo-Json -Compress)
                        }
                        else {
                            $flatRecord | Add-Member -MemberType NoteProperty -Name $property.Name -Value $value
                        }
                    }
                    
                    $flatRecord
                }
                
                $fileName = "CopilotUsage_${ReportType}_${timestamp}.csv"
                $filePath = Join-Path $ExportPath $fileName
                $flattenedData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
            }
        }
        
        Write-Log "Exported $($flattenedData.Count) records to: $filePath" -Level Success
        return $filePath
    }
    catch {
        Write-Log "Error exporting to CSV: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Upload-ToSharePoint {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$LibraryName,
        
        [Parameter(Mandatory = $false)]
        [string]$FolderPath = ""
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Log "File not found: $FilePath" -Level Error
            return $false
        }
        
        $fileName = Split-Path $FilePath -Leaf
        Write-Log "Uploading $fileName to SharePoint..."
        
        # Construct target folder path
        if ($FolderPath) {
            $targetFolder = "$LibraryName/$FolderPath"
        }
        else {
            $targetFolder = $LibraryName
        }
        
        # Upload file to SharePoint
        Add-PnPFile -Path $FilePath -Folder $targetFolder -ErrorAction Stop | Out-Null
        
        Write-Log "Successfully uploaded $fileName to SharePoint" -Level Success
        return $true
    }
    catch {
        Write-Log "Error uploading to SharePoint: $($_.Exception.Message)" -Level Error
        return $false
    }
}

#endregion

#region Main Script

try {
    Write-Log "=== Starting Copilot Usage Report Export ===" -Level Info
    
    # Load configuration
    if (-not (Test-Path $ConfigFile)) {
        Write-Log "Configuration file not found: $ConfigFile" -Level Error
        throw "Configuration file not found. Please create a config.json file."
    }
    
    $script:config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
    Write-Log "Configuration loaded from: $ConfigFile"
    
    # Validate required configuration
    $requiredProperties = @('TenantId', 'ClientId', 'CertificateThumbprint', 'SharePointSiteUrl', 'SharePointLibrary', 'LocalExportPath')
    foreach ($prop in $requiredProperties) {
        if (-not $script:config.$prop) {
            throw "Missing required configuration property: $prop"
        }
    }
    
    # Ensure LocalExportPath exists
    if (-not (Test-Path $script:config.LocalExportPath)) {
        Write-Log "Creating local export directory: $($script:config.LocalExportPath)"
        New-Item -Path $script:config.LocalExportPath -ItemType Directory -Force | Out-Null
    }
    
    # Load PnP.PowerShell module
    if ($script:config.'PnP.PowerShellModuleLocation' -and (Test-Path $script:config.'PnP.PowerShellModuleLocation')) {
        # Load module from local path
        Write-Log "Loading PnP.PowerShell module from local path: $($script:config.'PnP.PowerShellModuleLocation')"
        Import-Module $script:config.'PnP.PowerShellModuleLocation' -ErrorAction Stop
        Write-Log "PnP.PowerShell module loaded from local path" -Level Success
    }
    elseif ($script:config.'PnP.PowerShellModuleLocation') {
        # Path specified but doesn't exist
        Write-Log "PnP.PowerShell module path specified but not found: $($script:config.'PnP.PowerShellModuleLocation')" -Level Error
        throw "PnP.PowerShell module path specified in configuration does not exist. Please verify the path."
    }
    else {
        # No local path specified, check if module is installed
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            Write-Log "PnP.PowerShell module not found. Please install it using: Install-Module -Name PnP.PowerShell -Scope CurrentUser" -Level Error
            throw "Required module PnP.PowerShell is not installed. Please install it and try again."
        }
        
        # Import PnP.PowerShell module from installed location
        Import-Module PnP.PowerShell -ErrorAction Stop
        Write-Log "PnP.PowerShell module loaded"
    }
    
    # Connect to Microsoft Graph using certificate authentication
    Write-Log "Connecting to Microsoft Graph..."
    Connect-PnPOnline -Url $script:config.SharePointSiteUrl `
        -ClientId $script:config.ClientId `
        -Tenant $script:config.TenantId `
        -Thumbprint $script:config.CertificateThumbprint `
        -ErrorAction Stop
    
    Write-Log "Successfully connected to Microsoft Graph" -Level Success
    
    # Get the report types from config (or use defaults)
    $reportTypes = if ($script:config.ReportTypes) { $script:config.ReportTypes } else { @('UserDetail', 'ActivityUserDetail') }
    $period = if ($script:config.ReportPeriodDays) { $script:config.ReportPeriodDays } else { 7 }
    
    # Array to store exported file paths
    $exportedFiles = @()
    
    # Process each report type
    foreach ($reportType in $reportTypes) {
        Write-Log "Processing report: $reportType"
        
        # Retrieve report data
        $reportData = Get-CopilotUsageReport -ReportType $reportType -Period $period
        
        if ($reportData -and $reportData.Count -gt 0) {
            # Export to CSV
            $csvPath = Export-ToCSV -Data $reportData -ReportType $reportType -ExportPath $script:config.LocalExportPath
            
            if ($csvPath) {
                $exportedFiles += $csvPath
            }
        }
    }
    
    # Upload files to SharePoint
    if ($exportedFiles.Count -gt 0) {
        Write-Log "Uploading $($exportedFiles.Count) file(s) to SharePoint..."
        
        $folderPath = if ($script:config.SharePointFolder) { $script:config.SharePointFolder } else { "" }
        
        foreach ($file in $exportedFiles) {
            Upload-ToSharePoint -FilePath $file `
                -SiteUrl $script:config.SharePointSiteUrl `
                -LibraryName $script:config.SharePointLibrary `
                -FolderPath $folderPath
        }
    }
    else {
        Write-Log "No files to upload to SharePoint" -Level Warning
    }
    
    # Clean up old local files if retention days is configured
    if ($script:config.LocalRetentionDays -and $script:config.LocalRetentionDays -gt 0) {
        Write-Log "Cleaning up local files older than $($script:config.LocalRetentionDays) days..."
        $cutoffDate = (Get-Date).AddDays(-$script:config.LocalRetentionDays)
        Get-ChildItem -Path $script:config.LocalExportPath -Filter "CopilotUsage_*.csv" | 
        Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
        ForEach-Object {
            Remove-Item $_.FullName -Force
            Write-Log "Deleted old file: $($_.Name)"
        }
    }
    
    Write-Log "=== Copilot Usage Report Export Completed Successfully ===" -Level Success
}
catch {
    Write-Log "Critical error: $($_.Exception.Message)" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
    exit 1
}
finally {
    # Disconnect from PnP
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Log "Disconnected from PnP Online"
    }
    catch {
        # Ignore disconnect errors
    }
}

#endregion
