# Step 1: Check if the SharePoint Online Management Shell module is installed
# If not, install it from the PowerShell Gallery
if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
    Write-Host "SharePoint Online Management Shell module not found. Installing..."
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
} else {
    Write-Host "SharePoint Online Management Shell module is already installed."
}

# Step 2: Update the module to ensure you're using the latest version
Write-Host "Updating SharePoint Online Management Shell module..."
Update-Module -Name Microsoft.Online.SharePoint.PowerShell
Write-Host "Module update complete."

# Step 3: Import the module
Import-Module Microsoft.Online.SharePoint.PowerShell
Write-Host "Module imported successfully."

# Step 4: Connect to SharePoint Online using MFA
# This will prompt for interactive login with MFA
Write-Host "Connecting to SharePoint Online using MFA..."
Connect-SPOService
Write-Host "Connected to SharePoint Online successfully."

# Step 5: Enable Azure Information Protection (AIP) integration
# This allows sensitivity labels to be used in SharePoint and OneDrive
Write-Host "Enabling Azure Information Protection (AIP) integration..."
Set-SPOTenant -EnableAIPIntegration $true
Write-Host "AIP integration enabled successfully."
