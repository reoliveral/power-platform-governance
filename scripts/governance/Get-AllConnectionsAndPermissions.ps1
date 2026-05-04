# Get-AllConnectionsAndPermissions.ps1
# Lists all Power Platform connections by environment and their role assignments (permissions)
# Requires: Microsoft.PowerApps.Administration.PowerShell module
# Install: Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Force

#Requires -Modules Microsoft.PowerApps.Administration.PowerShell

param(
    [string]$OutputCsvPath = "",   # Optional: export results to CSV
    [string[]]$EnvironmentFilter = @()  # Optional: filter by specific environment name(s) or ID(s)
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

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

# ── Iterate environments ───────────────────────────────────────────────────────
foreach ($env in $environments) {
    $envName    = $env.EnvironmentName
    $envDisplay = $env.DisplayName
    $envRegion  = $env.Location

    Write-Host "Environment: $envDisplay ($envName)" -ForegroundColor Yellow

    # Get all connections in this environment
    try {
        $connections = Get-AdminPowerAppConnection -EnvironmentName $envName -ErrorAction Stop
    }
    catch {
        Write-Warning "  Could not retrieve connections for '$envDisplay': $_"
        continue
    }

    if (-not $connections) {
        Write-Host "  No connections found." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Found $($connections.Count) connection(s)." -ForegroundColor Green

    foreach ($conn in $connections) {
        $connId          = $conn.ConnectionName
        $connDisplayName = $conn.DisplayName
        $connectorName   = $conn.ConnectorName
        $connStatus      = $conn.Statuses[0].status
        $connCreatedBy   = $conn.CreatedBy.userPrincipalName

        Write-Host "    Connection: $connDisplayName [$connId]" -ForegroundColor White

        # Get role assignments (permissions) for this connection
        try {
            $roleAssignments = Get-AdminPowerAppConnectionRoleAssignment `
                -EnvironmentName $envName `
                -ConnectionName  $connId `
                -ConnectorName $connectorName 
        }
        catch {
            Write-Warning "      Could not retrieve permissions for connection '$connDisplayName': $_"
            $roleAssignments = @()
        }

        if (-not $roleAssignments) {
            # Record the connection even when there are no explicit role assignments
            $results.Add([PSCustomObject]@{
                EnvironmentId       = $envName
                EnvironmentName     = $envDisplay
                EnvironmentRegion   = $envRegion
                ConnectionId        = $connId
                ConnectionName      = $connDisplayName
                ConnectorName       = $connectorName
                ConnectionStatus    = $connStatus
                ConnectionCreatedBy = $connCreatedBy
                PrincipalType       = "N/A"
                PrincipalId         = "N/A"
                PrincipalEmail      = "N/A"
                RoleName            = "N/A"
            })
            continue
        }

        foreach ($role in $roleAssignments) {
            $results.Add([PSCustomObject]@{
                EnvironmentId       = $envName
                EnvironmentName     = $envDisplay
                EnvironmentRegion   = $envRegion
                ConnectionId        = $connId
                ConnectionName      = $connDisplayName
                ConnectorName       = $connectorName
                ConnectionStatus    = $connStatus
                ConnectionCreatedBy = $connCreatedBy
                PrincipalType       = $role.PrincipalType
                PrincipalId         = $role.PrincipalObjectId
                PrincipalEmail      = $role.PrincipalEmail
                RoleName            = $role.RoleName
            })

            $principal = if ($role.PrincipalEmail) { $role.PrincipalEmail } else { $role.PrincipalObjectId }
            Write-Host "      $($role.RoleName) -> $principal ($($role.PrincipalType))" `
                -ForegroundColor DarkCyan
        }
    }

    Write-Host ""
}

# ── Output ─────────────────────────────────────────────────────────────────────
Write-Host "`n=== Summary: $($results.Count) record(s) ===`n" -ForegroundColor Cyan
$results | Format-Table -AutoSize

if ($OutputCsvPath) {
    $dir = Split-Path -Path $OutputCsvPath -Parent
    if ($dir -and -not (Test-Path -Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $results | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $OutputCsvPath" -ForegroundColor Green
}
