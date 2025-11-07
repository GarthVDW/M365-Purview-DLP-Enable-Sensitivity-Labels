# Step 1: Check if the Exchange Online Management module is installed
# This module includes the Connect-IPPSSession cmdlet needed to access the Security & Compliance Center
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing..."
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
} else {
    Write-Host "ExchangeOnlineManagement module is already installed."
}

# Step 2: Import the module to make its cmdlets available in the session
Import-Module ExchangeOnlineManagement
Write-Host "ExchangeOnlineManagement module imported successfully."

# Step 3: Connect to the Security & Compliance Center using MFA
# This will prompt the user to sign in interactively with MFA
Write-Host "Connecting to Security & Compliance Center using MFA..."
Connect-IPPSSession

# Step 4: Execute the label sync command
# This synchronizes sensitivity labels from Microsoft Purview into Azure AD
Write-Host "Executing Azure AD Label Sync..."
Execute-AzureADLabelSync
Write-Host "Label sync command executed. Changes may take up to 24 hours to reflect."