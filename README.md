# Copilot Export Usage Reports

Automated PowerShell solution to export Microsoft 365 Copilot usage reports from Graph API and upload them to SharePoint Online.

## Features

- **Certificate-based Authentication**: Secure authentication using Azure AD App Registration with certificate
- **Multiple Report Types**: Supports various Copilot usage report types
- **Automated CSV Export**: Exports reports to CSV files locally
- **SharePoint Upload**: Automatically uploads reports to SharePoint Online
- **Scheduled Execution**: Can be configured to run daily via Windows Task Scheduler
- **Logging**: Comprehensive logging for auditing and troubleshooting
- **Retention Management**: Automatically cleans up old local files

## Prerequisites

1. **PowerShell 5.1 or later** (PowerShell 7+ recommended)
2. **PnP.PowerShell module** (must be installed manually - see installation instructions below)
3. **Azure AD App Registration** with:
   - Certificate-based authentication configured
   - Required Microsoft Graph API permissions:
     - `Reports.Read.All` (Application permission)
     - `Sites.Selected` (Application permission for SharePoint upload)
4. **SharePoint Online site** with a document library for storing reports and appropriate site permissions granted to the app

## Setup Instructions

### 1. Install PnP.PowerShell Module

Before running the script, you must install the PnP.PowerShell module:

```powershell
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

If you prefer to install for all users (requires admin):

```powershell
Install-Module -Name PnP.PowerShell -Scope AllUsers -Force
```

Verify the installation:

```powershell
Get-Module -ListAvailable -Name PnP.PowerShell
```

### 2. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App Registrations
2. Create a new app registration or use an existing one
3. Note the **Application (client) ID** and **Directory (tenant) ID**

### 3. Configure Certificate Authentication

You can either use an existing certificate or create a new one:

#### Option A: Create a Self-Signed Certificate

```powershell
# Create a self-signed certificate
$cert = New-SelfSignedCertificate -Subject "CN=CopilotUsageReportsApp" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export certificate (without private key) for Azure
$certPath = "C:\Temp\CopilotUsageReportsApp.cer"
Export-Certificate -Cert $cert -FilePath $certPath

# Note the thumbprint
Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
```

#### Option B: Use an Existing Certificate

Ensure the certificate is installed in your certificate store and note its thumbprint.

### 4. Upload Certificate to Azure AD App

1. In your App Registration, go to **Certificates & secrets**
2. Click **Upload certificate**
3. Upload the .cer file created above
4. Copy the **Thumbprint** value

### 5. Configure API Permissions

1. In your App Registration, go to **API permissions**
2. Add the following **Application permissions**:
   - **Microsoft Graph** > `Reports.Read.All`
   - **SharePoint** > `Sites.Selected`
3. Click **Grant admin consent** for your tenant

### 6. Grant Site Permissions to the App

Since we're using `Sites.Selected`, you need to explicitly grant the app access to the specific SharePoint site:

```powershell
# Connect to your SharePoint site
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/YourSite" -Interactive

# Grant write permissions to the app (use the ClientId from your app registration)
$clientId = "your-app-registration-client-id"
Grant-PnPAzureADAppSitePermission -AppId $clientId -DisplayName "Copilot Reports App" -Site (Get-PnPSite) -Permissions Write

# Verify the permission was granted
Get-PnPAzureADAppSitePermission
```

Alternatively, you can use Microsoft Graph PowerShell:

```powershell
Connect-MgGraph -Scopes "Sites.FullControl.All"

# Get the site ID
$siteUrl = "https://yourtenant.sharepoint.com/sites/YourSite"
$site = Get-MgSite -SiteId "yourtenant.sharepoint.com:/sites/YourSite"

# Grant write permission
$params = @{
    roles = @("write")
    grantedToIdentities = @(
        @{
            application = @{
                id = "your-app-registration-client-id"
                displayName = "Copilot Reports App"
            }
        }
    )
}
New-MgSitePermission -SiteId $site.Id -BodyParameter $params
```

### 7. Configure the Script

1. Edit `config.json` with your environment details:

```json
{
  "TenantId": "yourtenant.onmicrosoft.com",
  "ClientId": "12345678-1234-1234-1234-123456789abc",
  "CertificateThumbprint": "A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6Q7R8S9T0",
  "SharePointSiteUrl": "https://yourtenant.sharepoint.com/sites/YourSite",
  "SharePointLibrary": "Shared Documents",
  "SharePointFolder": "CopilotReports",
  "LocalExportPath": "C:\\CopilotReports\\Exports",
  "ReportTypes": [
    "UserDetail",
    "UserCountsSummary",
    "UserCountsTrend"
  ],
  "ReportPeriodDays": 7,
  "LocalRetentionDays": 30
}
```

### Configuration Options

| Property | Description | Required |
|----------|-------------|----------|
| `TenantId` | Your Azure AD tenant ID or domain | Yes |
| `ClientId` | Application (client) ID from app registration | Yes |
| `CertificateThumbprint` | Thumbprint of the certificate | Yes |
| `SharePointSiteUrl` | Full URL of your SharePoint site | Yes |
| `SharePointLibrary` | Name of the document library | Yes |
| `SharePointFolder` | Subfolder within the library (optional) | No |
| `LocalExportPath` | Local directory for CSV exports | Yes |
| `ReportTypes` | Array of report types to export | No |
| `ReportPeriodDays` | Number of days of data to retrieve (7, 30, 90, 180) | No |
| `LocalRetentionDays` | Days to retain local CSV files (0 = keep forever) | No |

### Available Report Types

- **UserDetail**: Detailed user-level Copilot usage
- **ActivityUserDetail**: User activity details
- **ActivityCounts**: Aggregated activity counts
- **ActivityUserCounts**: User count metrics

## Usage

### Manual Execution

Run the script manually:

```powershell
.\Export-CopilotUsageReports.ps1
```

With a custom config file:

```powershell
.\Export-CopilotUsageReports.ps1 -ConfigFile "C:\Path\To\config.json"
```

### Scheduled Execution

Use the included setup script to create a daily scheduled task:

```powershell
.\Setup-ScheduledTask.ps1
```

Or manually create a scheduled task:

1. Open Task Scheduler
2. Create a new task with the following action:
   - Program: `powershell.exe`
   - Arguments: `-ExecutionPolicy Bypass -File "C:\Path\To\Export-CopilotUsageReports.ps1"`
3. Set trigger to run daily at your preferred time
4. Configure the task to run whether user is logged on or not

## Output

### CSV Files

CSV files are saved to the configured `LocalExportPath` with the following naming convention:

```
CopilotUsage_<ReportType>_<Timestamp>.csv
```

Example: `CopilotUsage_UserDetail_20260105_143022.csv`

### Log Files

Log files are created daily in the same directory as the CSV exports:

```
CopilotExport_<Date>.log
```

Example: `CopilotExport_20260105.log`

### SharePoint Upload

Files are uploaded to:

```
<SharePointSiteUrl>/<SharePointLibrary>/<SharePointFolder>/
```

## Troubleshooting

### Authentication Errors

- Verify the certificate is installed in the correct certificate store
- Ensure the certificate thumbprint matches exactly (no spaces)
- Check that admin consent has been granted for API permissions

### Permission Errors

- Verify the app has `Reports.Read.All` and `Sites.Selected` permissions
- Ensure admin consent has been granted
- For SharePoint access, verify the app has been granted site permissions using `Grant-PnPAzureADAppSitePermission` or Microsoft Graph
- Check permissions with: `Get-PnPAzureADAppSitePermission` when connected to the site
- Ensure the app has at least **Write** permissions to the target SharePoint site

### Module Errors

If you haven't installed the PnP.PowerShell module, you'll see an error:

```
Required module PnP.PowerShell is not installed. Please install it and try again.
```

Install the module manually:

```powershell
Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
```

### SharePoint Upload Errors

- Verify the SharePoint site URL is correct and accessible
- Check that the document library exists
- Ensure the app has write permissions to the SharePoint site
- If using a folder, verify it exists or the script will try to create it

## Security Considerations

1. **Certificate Security**: Store certificates securely and limit access
2. **Config File**: Protect the config.json file - it contains sensitive information
3. **Service Account**: Consider running the scheduled task under a dedicated service account
4. **Permissions**: Follow the principle of least privilege for API permissions

## License

This project is provided as-is for use within Microsoft 365 environments.

## Support

For issues or questions, please refer to:
- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [Microsoft Graph API Documentation](https://docs.microsoft.com/graph/)
