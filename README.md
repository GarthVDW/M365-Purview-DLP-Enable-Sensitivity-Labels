# Microsoft 365 Purview DLP - Enable Sensitivity Labels

Complete PowerShell automation toolkit for enabling and configuring Microsoft Information Protection (MIP) sensitivity labels across your Microsoft 365 tenant.

## Overview

This repository contains PowerShell scripts that automate the complete setup process for sensitivity labels in Microsoft 365, including configuration for Groups, SharePoint Sites, OneDrive, and Azure AD synchronization.

## Features

- ✅ Automated module installation and updates (Microsoft Graph, Exchange Online, SharePoint Online)
- ✅ Microsoft Graph directory settings configuration for Groups and Sites
- ✅ SharePoint Online Azure Information Protection (AIP) integration
- ✅ Security & Compliance Center label synchronization
- ✅ Comprehensive error handling and logging
- ✅ Interactive authentication with MFA support
- ✅ Idempotent operations (safe to run multiple times)
- ✅ Detailed progress reporting and success verification

## Scripts Included

### Enable-SensitivityLabels-Complete.ps1
The comprehensive script that handles the complete sensitivity labels setup process across all Microsoft 365 services.

**What it does:**
1. Installs/updates required PowerShell modules
2. Connects to Microsoft Graph with appropriate permissions
3. Enables MIP labels for Microsoft 365 Groups
4. Enables MIP labels for SharePoint Sites
5. Enables SharePoint Online AIP integration
6. Connects to Security & Compliance Center
7. Syncs sensitivity labels to Azure AD

## Prerequisites

- PowerShell 5.1 or later
- Global Administrator or Compliance Administrator role
- Microsoft 365 tenant with appropriate licenses
- Internet connection for module installation and Microsoft 365 connectivity

## Required Permissions

- `Directory.ReadWrite.All` (Microsoft Graph)
- Exchange Online administrative access
- SharePoint Online administrative access
- Security & Compliance Center access

## Installation

1. Clone this repository:
   ```powershell
   git clone https://github.com/GarthVDW/M365-Purview-DLP-Enable-Sensitivity-Labels.git
   cd M365-Purview-DLP-Enable-Sensitivity-Labels
   ```

2. Ensure your execution policy allows running scripts:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage

### Basic Usage
Run the complete setup script with default settings:
```powershell
.\Enable-SensitivityLabels-Complete.ps1
```

### Skip Module Installation
If modules are already installed:
```powershell
.\Enable-SensitivityLabels-Complete.ps1 -SkipModuleInstall
```

### Custom Log Path
Specify a custom location for log files:
```powershell
.\Enable-SensitivityLabels-Complete.ps1 -LogPath "C:\Logs\M365"
```

### Force Module Reinstall
Force reinstallation of all modules:
```powershell
.\Enable-SensitivityLabels-Complete.ps1 -Force
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `LogPath` | String | No | Path for transcript logs (default: current directory) |
| `SkipModuleInstall` | Switch | No | Skip module installation if already installed |
| `Force` | Switch | No | Force reinstall modules even if present |

## What Gets Configured

1. **Microsoft Graph Directory Settings**
   - Enables `EnableMIPLabels` setting in Group.Unified template
   - Allows sensitivity labels on Microsoft 365 Groups

2. **SharePoint Online**
   - Enables Azure Information Protection integration
   - Allows sensitivity labels on SharePoint Sites and OneDrive

3. **Azure AD Label Sync**
   - Synchronizes sensitivity labels from Microsoft Purview to Azure AD
   - Makes labels available across all Microsoft 365 services

## Logging

All script executions create detailed transcript logs with timestamps:
- Format: `EnableSensitivityLabels_Complete_YYYYMMDD-HHMMSS.log`
- Location: Specified by `-LogPath` parameter or current directory
- Contains: All commands, output, errors, and stack traces

## Post-Configuration Steps

After running the script:

1. **Wait for propagation** - Changes may take up to 24 hours to reflect across all services
2. **Create sensitivity labels** in Microsoft Purview compliance portal
3. **Publish labels** to appropriate users and groups
4. **Test label application** on Teams, Groups, and SharePoint Sites
5. **Verify labels** appear in SharePoint and OneDrive

## Troubleshooting

### "Function capacity exceeded" Error
The script uses targeted Microsoft Graph modules to avoid this issue. If you still encounter it, try:
```powershell
# Close and reopen PowerShell, then run:
.\Enable-SensitivityLabels-Complete.ps1 -SkipModuleInstall
```

### Authentication Issues
If you encounter authentication errors:
1. Ensure you have the required admin roles
2. Try closing and reopening PowerShell as Administrator
3. Clear cached credentials: `Disconnect-MgGraph` and `Disconnect-ExchangeOnline`

### Labels Not Appearing
- Wait up to 24 hours for full propagation
- Verify labels are published to your users
- Check that users are licensed appropriately

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - feel free to use and modify for your organization's needs.

## Author

Garth van der Woude  
Created: November 7, 2025

## Disclaimer

This script is provided as-is with no warranties. Always test in a non-production environment first. Ensure you have proper backups and understand the changes being made to your Microsoft 365 tenant.

## Support

For issues, questions, or contributions, please open an issue in this repository.
