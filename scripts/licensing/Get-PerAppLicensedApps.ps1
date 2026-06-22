#Requires -Version 5.1

<#
.SYNOPSIS
    Gets Power Apps that are using per-app licensing in Power Platform.

.DESCRIPTION
    This script retrieves all Power Apps configured to use per-app licensing
    across all environments (or a specified environment) in Power Platform.
    Requires the Microsoft.PowerApps.Administration.PowerShell module.

    Per-app licensing allows users to run specific apps without needing a
    full Power Apps per-user license. This script identifies apps designated
    for per-app plan consumption.

.PARAMETER EnvironmentName
    Optional. The environment name (GUID) to scope the search to a specific environment.
    If omitted, all environments are scanned.

.PARAMETER ExportPath
    Optional. Full file path to export results as a CSV file.
    Example: "C:\Reports\PerAppApps.csv"

.PARAMETER ShowAllApps
    Optional. When specified, outputs all apps with their license designation,
    not just those marked as per-app. Useful for auditing.

.PARAMETER Diagnose
    Optional. Dumps the raw Internal.properties of the first app found in each environment.
    Use this to discover the exact property names returned by your tenant's API version
    when the script returns no results.

.EXAMPLE
    .\Get-PerAppLicensedApps.ps1

.EXAMPLE
    .\Get-PerAppLicensedApps.ps1 -EnvironmentName "00000000-0000-0000-0000-000000000000"

.EXAMPLE
    .\Get-PerAppLicensedApps.ps1 -ExportPath "C:\Reports\PerAppApps.csv"

.EXAMPLE
    .\Get-PerAppLicensedApps.ps1 -ShowAllApps -ExportPath "C:\Reports\AllAppsWithLicense.csv"

.EXAMPLE
    .\Get-PerAppLicensedApps.ps1 -Diagnose
    # Dumps raw app properties to identify the correct license property name in your tenant

.NOTES
    Required module: Microsoft.PowerApps.Administration.PowerShell
    Install with: Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$EnvironmentName,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        $parent = Split-Path $_ -Parent
        if ($parent -and -not (Test-Path $parent)) {
            throw "Export directory does not exist: $parent"
        }
        $true
    })]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [switch]$ShowAllApps,

    [Parameter(Mandatory = $false)]
    [switch]$Diagnose
)

$ErrorActionPreference = 'Stop'

#region Helper Functions

function Test-PowerAppsAdminModule {
    $moduleName = 'Microsoft.PowerApps.Administration.PowerShell'
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Warning "Module '$moduleName' is not installed."
        $install = Read-Host "Install it now for the current user? (Y/N)"
        if ($install -eq 'Y') {
            Write-Host "Installing $moduleName..." -ForegroundColor Yellow
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
        }
        else {
            throw "Required module '$moduleName' is not installed. Exiting."
        }
    }
    Import-Module -Name $moduleName -ErrorAction Stop
    Write-Host "Module '$moduleName' loaded." -ForegroundColor Green
}

function Connect-ToPowerPlatform {
    try {
        # Check if already authenticated by calling a lightweight admin cmdlet
        $null = Get-AdminPowerAppEnvironment -ErrorAction Stop -Top 1
        Write-Host "Already authenticated to Power Platform." -ForegroundColor Green
    }
    catch {
        Write-Host "Authenticating to Power Platform..." -ForegroundColor Yellow
        Add-PowerAppsAccount
    }
}

function Get-LicenseDesignation {
    param ($AppProperties)

    # Try all known property paths — the exact name varies by tenant/API version.
    # Run the script with -Diagnose to see which properties your tenant returns.
    $candidateProperties = @(
        'licenseDesignation',
        'planClassification',
        'licenseType',
        'appLicenseType',
        'appPlanClassification'
    )

    foreach ($prop in $candidateProperties) {
        $value = $AppProperties.$prop
        if ($value) { return $value }
    }

    # Boolean fallbacks
    if ($AppProperties.premiumRequired -eq $true) { return 'Premium' }
    if ($AppProperties.bypassConsent   -eq $true) { return 'Premium' }

    return 'Standard'
}

function Resolve-IsPerApp {
    param ([string]$LicenseDesignation)
    return ($LicenseDesignation -in @('PerApp', 'perApp', 'Per App', 'per_app', 'PerAppPlan'))
}

function Write-AppPropertiesDiagnostic {
    param (
        $App,
        [string]$EnvironmentDisplay
    )
    Write-Host "`n[DIAGNOSTIC] App: '$($App.DisplayName)' in '$EnvironmentDisplay'" -ForegroundColor Magenta
    Write-Host "--- Internal.properties members ---" -ForegroundColor Magenta
    $App.Internal.properties | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name | ForEach-Object {
        $val = $App.Internal.properties.$_
        Write-Host "  $_ = $val" -ForegroundColor Gray
    }
    Write-Host "--- app-level Tags ---" -ForegroundColor Magenta
    $App.Tags | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | ForEach-Object {
        Write-Host "  [Tag] $_ = $($App.Tags.$_)" -ForegroundColor Gray
    }
}

#endregion

#region Main

Write-Host "`n=== Power Platform Per-App Licensed Apps Report ===" -ForegroundColor Cyan
Write-Host "Run date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor Gray

# Load required module
Test-PowerAppsAdminModule

# Authenticate
Connect-ToPowerPlatform

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

# Retrieve environments
if ($EnvironmentName) {
    Write-Host "Fetching environment: $EnvironmentName" -ForegroundColor Cyan
    try {
        $environments = @(Get-AdminPowerAppEnvironment -EnvironmentName $EnvironmentName)
    }
    catch {
        throw "Could not retrieve environment '$EnvironmentName'. Verify the GUID and your permissions. Error: $_"
    }
}
else {
    Write-Host "Fetching all environments..." -ForegroundColor Cyan
    $environments = @(Get-AdminPowerAppEnvironment)
}

if ($environments.Count -eq 0) {
    Write-Warning "No environments found. Verify your account has Power Platform admin permissions."
    return
}

Write-Host "Found $($environments.Count) environment(s).`n" -ForegroundColor Green

$envIndex = 0
foreach ($env in $environments) {
    $envIndex++
    $envDisplay = "$($env.DisplayName) [$($env.EnvironmentName)]"
    Write-Host "[$envIndex/$($environments.Count)] Processing: $envDisplay" -ForegroundColor Cyan

    try {
        $apps = @(Get-AdminPowerApp -EnvironmentName $env.EnvironmentName)
        Write-Host "  Found $($apps.Count) app(s)." -ForegroundColor Gray

        # Diagnostic: dump first app's properties to identify correct property names
        if ($Diagnose -and $apps.Count -gt 0) {
            Write-AppPropertiesDiagnostic -App $apps[0] -EnvironmentDisplay $env.DisplayName
        }

        foreach ($app in $apps) {
            $props             = $app.Internal.properties
            $licenseDesig      = Get-LicenseDesignation -AppProperties $props
            $isPerApp          = Resolve-IsPerApp -LicenseDesignation $licenseDesig

            if ($isPerApp -or $ShowAllApps) {
                $results.Add([PSCustomObject]@{
                    AppName            = $app.AppName
                    DisplayName        = $app.DisplayName
                    EnvironmentName    = $env.EnvironmentName
                    EnvironmentDisplay = $env.DisplayName
                    OwnerUPN           = $app.Owner.userPrincipalName
                    OwnerObjectId      = $app.Owner.id
                    CreatedTime        = $app.CreatedTime
                    LastModifiedTime   = $app.LastModifiedTime
                    LicenseDesignation = $licenseDesig
                    IsPerApp           = $isPerApp
                    AppStatus          = $props.status
                    AppType            = $props.appType
                    SharedWithTenants  = $props.sharedWithTenantCount
                    SharedWithUsers    = $props.sharedWithUsersCount
                })
            }
        }
    }
    catch {
        Write-Warning "  Failed to retrieve apps for '$($env.DisplayName)': $_"
    }
}

# Display results
Write-Host "`n--- Results ---`n" -ForegroundColor Cyan

if ($results.Count -eq 0) {
    $scope = if ($ShowAllApps) { 'apps' } else { 'apps using per-app licensing' }
    Write-Host "No $scope were found." -ForegroundColor Yellow
    if (-not $ShowAllApps -and -not $Diagnose) {
        Write-Host "Tip: Run with -Diagnose to dump raw app properties and identify the correct license property name for your tenant." -ForegroundColor DarkYellow
        Write-Host "Tip: Run with -ShowAllApps to list all apps with their detected license designation." -ForegroundColor DarkYellow
    }
}
else {
    $label = if ($ShowAllApps) { 'app(s) found' } else { 'app(s) using per-app licensing' }
    Write-Host "Total: $($results.Count) $label" -ForegroundColor Green

    $results | Format-Table -Property DisplayName, EnvironmentDisplay, OwnerUPN, LicenseDesignation, AppStatus -AutoSize

    if ($ExportPath) {
        $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $ExportPath" -ForegroundColor Green
    }
}

# Return results for pipeline use
return $results

#endregion
