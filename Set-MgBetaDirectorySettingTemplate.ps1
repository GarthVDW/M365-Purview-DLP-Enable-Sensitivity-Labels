<#
.SYNOPSIS
    Enable Sensitivity Labels for Microsoft 365 Groups and SharePoint Sites

.DESCRIPTION
    This script automates the process of enabling Microsoft Information Protection (MIP) 
    sensitivity labels for Microsoft 365 Groups and SharePoint Sites. It handles module 
    installation, configuration, and label synchronization with comprehensive error handling.

.PARAMETER LogPath
    Path where the transcript log will be saved. Default: Current directory

.PARAMETER SkipModuleInstall
    Skip the module installation step if modules are already installed

.PARAMETER Force
    Force reinstall modules even if they already exist

.EXAMPLE
    .\EnableSensitivityLabels.ps1
    Run with default settings

.EXAMPLE
    .\EnableSensitivityLabels.ps1 -SkipModuleInstall -LogPath "C:\Logs"
    Skip module installation and save logs to specific path

.NOTES
    Author: Improved Script
    Date: 2025-11-07
    Requires: PowerShell 5.1 or later, Global Administrator or Compliance Administrator role
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".",
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipModuleInstall,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

#Requires -Version 5.1

# Initialize transcript logging
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$transcriptPath = Join-Path $LogPath "EnableSensitivityLabels_$timestamp.log"
Start-Transcript -Path $transcriptPath -Append

try {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Enable Sensitivity Labels for M365" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # Check PowerShell version
    Write-Verbose "Checking PowerShell version..."
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "This script requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
    }
    Write-Host "[OK] PowerShell version check passed" -ForegroundColor Green

    # Step 1: Install/Update required modules
    if (-not $SkipModuleInstall) {
        Write-Host "`n[1/7] Checking and installing required modules..." -ForegroundColor Yellow
        
        $requiredModules = @(
            @{Name = "Microsoft.Graph.Beta"; MinVersion = "2.0.0"},
            @{Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0"}
        )
        
        foreach ($module in $requiredModules) {
            Write-Host "  Checking $($module.Name)..." -NoNewline
            $installedModule = Get-Module -ListAvailable -Name $module.Name | 
                Sort-Object Version -Descending | 
                Select-Object -First 1
            
            if ($installedModule -and -not $Force) {
                Write-Host " [Installed: v$($installedModule.Version)]" -ForegroundColor Green
            } else {
                Write-Host " [Installing...]" -ForegroundColor Yellow
                try {
                    Install-Module $module.Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                    Write-Host "    [OK] $($module.Name) installed successfully" -ForegroundColor Green
                } catch {
                    throw "Failed to install $($module.Name): $_"
                }
            }
        }
    } else {
        Write-Host "`n[1/7] Skipping module installation (SkipModuleInstall specified)" -ForegroundColor Gray
    }

    # Step 2: Import required modules
    Write-Host "`n[2/7] Importing required modules..." -ForegroundColor Yellow
    try {
        Import-Module Microsoft.Graph.Beta -ErrorAction Stop
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-Host "  [OK] Modules imported successfully" -ForegroundColor Green
    } catch {
        throw "Failed to import modules: $_"
    }

    # Step 3: Connect to Microsoft Graph
    Write-Host "`n[3/7] Connecting to Microsoft Graph..." -ForegroundColor Yellow
    try {
        # Check if already connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-Host "  [INFO] Already connected as: $($context.Account)" -ForegroundColor Cyan
            $reconnect = Read-Host "  Do you want to reconnect? (Y/N)"
            if ($reconnect -eq 'Y') {
                Disconnect-MgGraph -ErrorAction SilentlyContinue
                Connect-MgGraph -Scopes "Directory.ReadWrite.All" -ErrorAction Stop
            }
        } else {
            Connect-MgGraph -Scopes "Directory.ReadWrite.All" -ErrorAction Stop
        }
        
        $context = Get-MgContext
        Write-Host "  [OK] Connected to Microsoft Graph" -ForegroundColor Green
        Write-Host "      Tenant: $($context.TenantId)" -ForegroundColor Gray
        Write-Host "      Account: $($context.Account)" -ForegroundColor Gray
    } catch {
        throw "Failed to connect to Microsoft Graph: $_"
    }

    # Step 4: Get Group.Unified template
    Write-Host "`n[4/7] Retrieving Group.Unified directory setting template..." -ForegroundColor Yellow
    try {
        $template = Get-MgBetaDirectorySettingTemplate | 
            Where-Object { $_.DisplayName -eq "Group.Unified" }
        
        if (-not $template) {
            throw "Group.Unified template not found in tenant"
        }
        Write-Host "  [OK] Template retrieved successfully (ID: $($template.Id))" -ForegroundColor Green
    } catch {
        throw "Failed to retrieve template: $_"
    }

    # Step 5: Check existing settings and create/update directory setting
    Write-Host "`n[5/7] Configuring MIP labels for Groups and Sites..." -ForegroundColor Yellow
    try {
        # Check if setting already exists
        $existingSetting = Get-MgBetaDirectorySetting | 
            Where-Object { $_.TemplateId -eq $template.Id }
        
        if ($existingSetting) {
            Write-Host "  [INFO] Directory setting already exists" -ForegroundColor Cyan
            
            # Check current EnableMIPLabels value
            $currentValue = $existingSetting.Values | 
                Where-Object { $_.Name -eq "EnableMIPLabels" } | 
                Select-Object -ExpandProperty Value
            
            if ($currentValue -eq "True") {
                Write-Host "  [OK] MIP labels already enabled - no changes needed" -ForegroundColor Green
            } else {
                Write-Host "  [INFO] Updating setting to enable MIP labels..." -ForegroundColor Yellow
                $values = $existingSetting.Values
                ($values | Where-Object { $_.Name -eq "EnableMIPLabels" }).Value = "True"
                
                Update-MgBetaDirectorySetting -DirectorySettingId $existingSetting.Id -Values $values -ErrorAction Stop
                Write-Host "  [OK] MIP labels enabled successfully" -ForegroundColor Green
            }
        } else {
            Write-Host "  [INFO] Creating new directory setting..." -ForegroundColor Yellow
            $setting = @{
                TemplateId = $template.Id
                Values = @(
                    @{ Name = "EnableMIPLabels"; Value = "True" }
                )
            }
            New-MgBetaDirectorySetting -BodyParameter $setting -ErrorAction Stop
            Write-Host "  [OK] MIP labels enabled successfully" -ForegroundColor Green
        }
    } catch {
        throw "Failed to configure directory setting: $_"
    }

    # Step 6: Connect to Security & Compliance Center
    Write-Host "`n[6/7] Connecting to Security & Compliance Center..." -ForegroundColor Yellow
    try {
        # Check if already connected
        $existingSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($existingSession) {
            Write-Host "  [INFO] Already connected to Security & Compliance Center" -ForegroundColor Cyan
        } else {
            Connect-IPPSSession -ErrorAction Stop
        }
        Write-Host "  [OK] Connected to Security & Compliance Center" -ForegroundColor Green
    } catch {
        throw "Failed to connect to Security & Compliance Center: $_"
    }

    # Step 7: Sync labels to Azure AD
    Write-Host "`n[7/7] Syncing sensitivity labels to Azure AD..." -ForegroundColor Yellow
    try {
        Execute-AzureAdLabelSync -ErrorAction Stop
        Write-Host "  [OK] Label sync completed successfully" -ForegroundColor Green
        Write-Host ""
        Write-Host "  [NOTE] It may take up to 24 hours for labels to appear in all services" -ForegroundColor Cyan
    } catch {
        throw "Failed to sync labels: $_"
    }

    # Success summary
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Configuration completed successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Cyan
    Write-Host "  1. Wait up to 24 hours for labels to propagate" -ForegroundColor White
    Write-Host "  2. Create sensitivity labels in Microsoft Purview" -ForegroundColor White
    Write-Host "  3. Publish labels to users and groups" -ForegroundColor White
    Write-Host "  4. Test label application on Teams/Groups/Sites" -ForegroundColor White
    Write-Host ""
    Write-Host "Log file saved to: $transcriptPath" -ForegroundColor Gray
    
    exit 0

} catch {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "ERROR: Script execution failed" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor Gray
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host ""
    Write-Host "Log file saved to: $transcriptPath" -ForegroundColor Gray
    
    exit 1
} finally {
    Stop-Transcript
}
