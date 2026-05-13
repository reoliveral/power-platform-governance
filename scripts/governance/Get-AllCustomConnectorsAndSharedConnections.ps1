# Get-AllCustomConnectorsAndSharedConnections.ps1
# Lists all Power Platform CUSTOM CONNECTORS and SHARED CONNECTIONS by environment,
# including their role assignments (who they are shared with and what role).
# Requires: Microsoft.PowerApps.Administration.PowerShell module
# Install: Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Force

#Requires -Modules Microsoft.PowerApps.Administration.PowerShell

param(
    [string]$OutputCsvPath = "",         # Optional: export results to CSV
    [string[]]$EnvironmentFilter = @()   # Optional: filter by environment name(s) or ID(s)
)

# ── Authentication ─────────────────────────────────────────────────────────────
# Comment out the line below if you are already signed in
Add-PowerAppsAccount

# ── Collect environments ───────────────────────────────────────────────────────
Write-Host "Retrieving environments..." -ForegroundColor Cyan
$environments = Get-AdminPowerAppEnvironment

if ($EnvironmentFilter.Count -gt 0) {
    $environments = $environments | Where-Object {
        $EnvironmentFilter -contains $_.EnvironmentName -or
        $EnvironmentFilter -contains $_.DisplayName
    }
}

Write-Host "Found $($environments.Count) environment(s).`n" -ForegroundColor Green

$customConnectorResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$sharedConnectionResults = [System.Collections.Generic.List[PSCustomObject]]::new()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — CUSTOM CONNECTORS & THEIR ROLE ASSIGNMENTS
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "══════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host " CUSTOM CONNECTORS" -ForegroundColor Magenta
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Magenta

foreach ($env in $environments) {
    $envName    = $env.EnvironmentName
    $envDisplay = $env.DisplayName
    $envRegion  = $env.Location

    Write-Host "Environment: $envDisplay ($envName)" -ForegroundColor Yellow

    # Get all custom connectors in this environment
    try {
        $customConnectors = Get-AdminPowerAppConnector -EnvironmentName $envName -ErrorAction Stop |
            Where-Object { $_.Internal.isCustomApi -eq $true }
    }
    catch {
        Write-Warning "  Could not retrieve custom connectors for '$envDisplay': $_"
        continue
    }

    if (-not $customConnectors) {
        Write-Host "  No custom connectors found." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Found $($customConnectors.Count) custom connector(s)." -ForegroundColor Green

    foreach ($connector in $customConnectors) {
        $connectorId          = $connector.ConnectorName
        $connectorDisplayName = $connector.DisplayName
        $connectorCreatedBy   = $connector.CreatedBy.userPrincipalName

        Write-Host "    Custom Connector: $connectorDisplayName [$connectorId]" -ForegroundColor White

        # Get role assignments (who the connector is shared with)
        try {
            $roleAssignments = Get-AdminPowerAppConnectorRoleAssignment `
                -EnvironmentName $envName `
                -ConnectorName   $connectorId `
                -ErrorAction Stop
        }
        catch {
            Write-Warning "      Could not retrieve permissions for connector '$connectorDisplayName': $_"
            $roleAssignments = @()
        }

        if (-not $roleAssignments) {
            $customConnectorResults.Add([PSCustomObject]@{
                EnvironmentId           = $envName
                EnvironmentName         = $envDisplay
                EnvironmentRegion       = $envRegion
                CustomConnectorId       = $connectorId
                CustomConnectorName     = $connectorDisplayName
                ConnectorCreatedBy      = $connectorCreatedBy
                SharedWithPrincipalType = "N/A"
                SharedWithPrincipalId   = "N/A"
                SharedWithEmail         = "N/A"
                RoleName                = "N/A (not shared)"
            })
            continue
        }

        foreach ($role in $roleAssignments) {
            $customConnectorResults.Add([PSCustomObject]@{
                EnvironmentId           = $envName
                EnvironmentName         = $envDisplay
                EnvironmentRegion       = $envRegion
                CustomConnectorId       = $connectorId
                CustomConnectorName     = $connectorDisplayName
                ConnectorCreatedBy      = $connectorCreatedBy
                SharedWithPrincipalType = $role.PrincipalType
                SharedWithPrincipalId   = $role.PrincipalObjectId
                SharedWithEmail         = $role.PrincipalEmail
                RoleName                = $role.RoleName
            })

            $principal = if ($role.PrincipalEmail) { $role.PrincipalEmail } else { $role.PrincipalObjectId }
            Write-Host "      $($role.RoleName) -> $principal ($($role.PrincipalType))" -ForegroundColor DarkCyan
        }
    }

    Write-Host ""
}

# ── Output ─────────────────────────────────────────────────────────────────────
Write-Host "`n=== Custom Connectors: $($customConnectorResults.Count) record(s) ===" -ForegroundColor Cyan
$customConnectorResults | Format-Table -AutoSize


# ── Optional CSV export ────────────────────────────────────────────────────────
if ($OutputCsvPath) {
    $dir = Split-Path -Path $OutputCsvPath -Parent
    if ($dir -and -not (Test-Path -Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $baseNoExt = [System.IO.Path]::GetFileNameWithoutExtension($OutputCsvPath)
    $ext       = [System.IO.Path]::GetExtension($OutputCsvPath)
    $folder    = if ($dir) { $dir } else { "." }

    $customCsvPath = Join-Path $folder "${baseNoExt}_CustomConnectors${ext}"

    $customConnectorResults  | Export-Csv -Path $customCsvPath -NoTypeInformation -Encoding UTF8
  
    Write-Host "Custom connectors exported to : $customCsvPath" -ForegroundColor Green
   }
