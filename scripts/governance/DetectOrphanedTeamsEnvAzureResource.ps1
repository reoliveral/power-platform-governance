
<#
.SYNOPSIS
    Identify Orphaned Dataverse for Teams Environments
.DESCRIPTION
This script connects to Azure, Microsoft Teams, and the Power Platform Administration API to 
identify orphaned Dataverse for Teams environments—those that remain in the tenant but are 
no longer linked to an active Microsoft Team or whose environment owner has left the company.
It is intended for Power Platform administrators and Center of Excellence (CoE) teams who need 
visibility and cleanup insights for Teams‑based Power Platform assets.
.EXAMPLE
    PS C:\> .\DetectOrphanedTeamsEnvAzureResource.ps1 
    Runs the script to check for orphaned Dataverse for Teams environments.
    The outcome will indicate whether orphaned environments were found or none were detected.
.INPUTS
    None
.OUTPUTS
    None
#>

# Install Azure Resource Graph module
Install-Module -Name Az.ResourceGraph -Repository PSGallery -Scope CurrentUser
Import-Module Az.ResourceGraph

# Install the MicrosoftTeams module (requires admin privileges)
Install-Module -Name MicrosoftTeams -Force -AllowClobber -Scope CurrentUser
Import-Module MicrosoftTeams


# Install the Power Platform Administration Module
Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
Import-Module Microsoft.PowerApps.Administration.PowerShell 

$WarningPreference = 'SilentlyContinue'

try {
    # Connect to Azure Subscription
    Write-Host "Establishing connection to Azure Subscription..." -ForegroundColor Blue
    $Credential = Get-Credential
    Connect-AzAccount -Credential $Credential -ErrorAction Stop -Verbose -WarningAction SilentlyContinue

    # Remove the comment if you have multiple subscriptions.
    # Set-AzContext -Subscription "<subscriptionguid>"
      
   # Apply a filter for Dataverse for Teams environments
    $teamsEnvironments = Search-AzGraph -Query "PowerPlatformResources | where type == 'microsoft.powerplatform/environments' | where properties.environmentType == 'Teams'"

   # Condition to check whether any Dataverse for Teams environments have been detected
    if (-not $teamsEnvironments) {
        Write-Host "No Dataverse for Teams environments found." -ForegroundColor Yellow
        return
    }

    # Connect to Microsoft Teams (interactive login)
    Write-Host "Establishing connection to Microsoft Teams..." -ForegroundColor Blue
    Connect-MicrosoftTeams -ErrorAction Stop -Verbose -WarningAction SilentlyContinue

    # Connect to Power Platform admin service
    Write-Host "Connecting to Power Platform..." -ForegroundColor Cyan
    Add-PowerAppsAccount -ErrorAction Stop -Verbose -WarningAction SilentlyContinue

    Write-Host "Checking for orphaned environments..." -ForegroundColor Cyan
    
    $orphanedEnvs = @()

    foreach ($env in $teamsEnvironments) {
        # Get details of Teams environment
        $teamsEnvDet = Get-AdminPowerAppEnvironment -EnvironmentName $env.Name
  
        # The Environment typically includes the associated Teams ID.
        $teamsId = $teamsEnvDet.Internal.properties.connectedGroups.id
        
        if ($teamsId -eq $null) {
            # No linked Teams ID means it's orphaned
            # The Teams group may be removed.
            $orphanedEnvs += $teamsEnvDet
            continue
        }

        # Try to get the linked Microsoft Teams team
        try {
            $team = Get-Team -GroupId $teamsId 

            # Get owners of the current team
            $owners = Get-TeamUser -GroupId $teamsId -Role Owner

            # Condition to verify when owners are not found
            if ($owners -eq $null) {
                   $orphanedEnvs += $teamsEnvDet
            }
        }
        catch {
            # If the team is unavailable or cannot be accessed, mark it as orphaned.
            $orphanedEnvs += $teamsEnvDet
        }
    }

    # Results
    if ($orphanedEnvs.Count -gt 0) {
        Write-Host "`nOrphaned Dataverse for Teams environments found:" -ForegroundColor Red
        $orphanedEnvs | Select-Object DisplayName, EnvironmentName, CreatedTime  | Format-Table
    }
    else {
        Write-Host "No orphaned environments detected." -ForegroundColor Green
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}