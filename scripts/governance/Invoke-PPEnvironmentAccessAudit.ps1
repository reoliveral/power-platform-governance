#Requires -Version 5.1

<#
.SYNOPSIS
    Power Platform - Single-Environment Access Auditor.

.DESCRIPTION
    Inventories every person, group, and service account that has access to
    resources in ONE Power Platform environment.

    Designed for the "one environment per run" workflow:
        1.  Run this script for environment A.
        2.  Review the CSVs with the business.
        3.  Create the Entra ID security group and bulk-import the seed file.
        4.  Associate the security group in the Power Platform admin centre.
        5.  Re-run this script for the next environment.

    Artefacts inventoried:
        * Canvas app owners and every role-assignment
          (users / groups / entire tenant)
        * Cloud flow creators and co-owners (users and groups)
        * Connection owners  (catches service accounts whose connections
          keep flows alive)
        * Custom connector owners
        * (-IncludeDataverse) Users holding at least one Dataverse
          security role
        * (-IncludeDataverse) Teams carrying security roles:
              Owner/Access-team members enumerated individually;
              Entra-group-backed teams reported as groups
        * (-IncludeDataverse) Copilot Studio agent owners and every
          principal each agent is shared with
        * (-IncludeDataverse) Solution cloud-flow owners
          (workflow table, category 5)
        * (-ExpandGroups) Transitive user membership of every group
          encountered, resolved via Microsoft Graph

    Output folder (one folder, seven CSVs):
        UserAccess_Detail.csv        - every artefact-level access record
        UserAccess_Summary.csv       - one row per principal + all reasons
        TenantWideShares.csv         - artefacts shared with the entire tenant
        GroupsEncountered.csv        - all groups referenced by access records
        GroupMembers.csv             - transitive user members (-ExpandGroups)
        UnresolvedPrincipals.csv     - object IDs that could not be resolved
        SecurityGroupSeed_<env>.csv  - Entra bulk-import-ready UPN list

    IMPORTANT - what this script does NOT prove:
    Owners and share targets are *intended* users, not *proven* usage.
    Before locking an environment down, cross-check against actual telemetry
    (Power Platform admin centre usage analytics, Dataverse audit logs, and
    the Microsoft 365 unified audit log for "Launched app" events).

.PARAMETER EnvironmentId
    The GUID of the Power Platform environment to audit.

.PARAMETER IncludeDataverse
    Also query the Dataverse instance: security-role holders, team
    memberships, Copilot Studio agent sharing, solution cloud-flow owners.

.PARAMETER ExpandGroups
    Resolve every discovered group to its transitive user members via Graph.
    Requires GroupMember.Read.All on the Graph connection.

.PARAMETER OutputFolder
    Destination folder for CSV files. Created if it does not exist.
    Default: .\PPAudit_<envName>_<yyyyMMdd_HHmmss>

.PARAMETER TenantId
    Optional. Azure / Entra tenant GUID.  Passed to Connect-AzAccount and
    Connect-MgGraph when the script handles authentication.

.EXAMPLE
    .\Invoke-PPEnvironmentAccessAudit.ps1 `
        -EnvironmentId "5f04367c-0b4d-e9cb-8685-f3a9ecc5cf12" `
        -IncludeDataverse `
        -ExpandGroups

.NOTES
    Required modules (install once):
        Install-Module Microsoft.PowerApps.Administration.PowerShell `
            -Scope CurrentUser -Force
        Install-Module Az.Accounts     -Scope CurrentUser -Force
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory, HelpMessage = 'Power Platform environment GUID')]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string] $EnvironmentId,

    [switch] $IncludeDataverse,
    [switch] $ExpandGroups,

    [string] $OutputFolder,

    [string] $TenantId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'Continue'

###############################################################################
# 0.  BANNER
###############################################################################
$scriptVersion = '1.0.0'
Write-Host ''
Write-Host '  +======================================================================+' -ForegroundColor Cyan
Write-Host "  |   Power Platform - Environment Access Auditor   v$scriptVersion              |" -ForegroundColor Cyan
Write-Host '  |   One environment per run.  Review > Group > Associate > Repeat.    |' -ForegroundColor Cyan
Write-Host '  +======================================================================+' -ForegroundColor Cyan
Write-Host ''

###############################################################################
# 1.  MODULE CHECK
###############################################################################
Write-Host '  -- Module check -------------------------------------------------------' -ForegroundColor Yellow

$requiredModules = @(
    'Microsoft.PowerApps.Administration.PowerShell'
)

# Accept the umbrella Microsoft.Graph OR any of the individual sub-modules
$graphUmbrella    = Get-Module -ListAvailable -Name 'Microsoft.Graph' -ErrorAction SilentlyContinue
$graphSubModules  = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups'
) | ForEach-Object { Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue }
$graphAvailable   = ($graphUmbrella -or ($graphSubModules | Where-Object { $_ }))

$missing = [System.Collections.Generic.List[string]]::new()
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
        $missing.Add($m)
    }
}
if (-not $graphAvailable) { $missing.Add('Microsoft.Graph') }

if ($missing.Count -gt 0) {
    Write-Host '  ERROR: Missing required modules:' -ForegroundColor Red
    foreach ($m in $missing) { Write-Host "    * $m" -ForegroundColor Red }
    Write-Host ''
    Write-Host '  Install them with:' -ForegroundColor Yellow
    Write-Host '    Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force' -ForegroundColor White
    Write-Host '    Install-Module Az.Accounts     -Scope CurrentUser -Force' -ForegroundColor White
    Write-Host '    Install-Module Microsoft.Graph -Scope CurrentUser -Force' -ForegroundColor White
    throw 'Missing required modules. Install them and re-run.'
}

Write-Host '  All required modules found.' -ForegroundColor Green
Write-Host ''

###############################################################################
# 2.  AUTHENTICATION
###############################################################################
Write-Host '  -- Authentication -----------------------------------------------------' -ForegroundColor Yellow

# -- 2a. Power Apps Administration ------------------------------------------
Write-Host '  [1/2] Power Apps Administration ...' -NoNewline -ForegroundColor White
$ppConnected = $false
try {
    $null = Get-AdminPowerAppEnvironment -EnvironmentName $EnvironmentId -ErrorAction Stop
    $ppConnected = $true
    Write-Host ' already connected.' -ForegroundColor Green
} catch { Write-Host '' }

if (-not $ppConnected) {
    Write-Host '        Launching interactive sign-in for Power Apps ...' -ForegroundColor Yellow
    if ($TenantId) {
        Add-PowerAppsAccount -TenantID $TenantId
    } else {
        Add-PowerAppsAccount
    }
    Write-Host '  [1/2] Power Apps Administration ... connected.' -ForegroundColor Green
}

# -- 2b. Microsoft Graph -----------------------------------------------------
Write-Host '  [2/2] Microsoft Graph ...' -NoNewline -ForegroundColor White
$graphScopes   = @('User.Read.All', 'Group.Read.All', 'GroupMember.Read.All')
$graphConnected = $false
try {
    $mgCtx = Get-MgContext -ErrorAction Stop
    if ($null -ne $mgCtx -and $null -ne $mgCtx.Account) {
        $missingScopes = $graphScopes | Where-Object { $mgCtx.Scopes -notcontains $_ }
        if ($missingScopes.Count -eq 0) {
            $graphConnected = $true
            Write-Host " already connected as $($mgCtx.Account)." -ForegroundColor Green
        } else {
            Write-Host " re-connecting (need: $($missingScopes -join ', ')) ..." -ForegroundColor Yellow
        }
    }
} catch { Write-Host '' }

if (-not $graphConnected) {
    Write-Host '        Launching interactive sign-in for Microsoft Graph ...' -ForegroundColor Yellow
    $mgParams = @{ Scopes = $graphScopes }
    if ($TenantId) { $mgParams['TenantId'] = $TenantId }
    try {
        Connect-MgGraph @mgParams -NoWelcome -ErrorAction Stop
    } catch {
        # -NoWelcome was added in a later module version - fall back gracefully
        Connect-MgGraph @mgParams -ErrorAction Stop
    }
    Write-Host '  [2/2] Microsoft Graph ... connected.' -ForegroundColor Green
}
Write-Host ''

###############################################################################
# 3.  ENVIRONMENT METADATA
###############################################################################
Write-Host '  -- Environment metadata -----------------------------------------------' -ForegroundColor Yellow

$envObj         = Get-AdminPowerAppEnvironment -EnvironmentName $EnvironmentId
$envDisplayName = $envObj.DisplayName
$instanceUrl    = $envObj.Internal.properties.linkedEnvironmentMetadata.instanceUrl
$envSafeName    = ($envDisplayName -replace '[^\w\-]', '_') -replace '_+', '_'
$envSafeName    = $envSafeName.Trim('_')

if (-not $OutputFolder) {
    $ts           = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFolder = Join-Path (Get-Location) "PPAudit_${envSafeName}_$ts"
}
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

Write-Host "  Environment   : $envDisplayName" -ForegroundColor White
Write-Host "  Environment ID: $EnvironmentId"  -ForegroundColor White
if ($instanceUrl) {
    Write-Host "  Dataverse URL : $instanceUrl" -ForegroundColor White
} else {
    Write-Host '  Dataverse URL : (none - not a Dataverse-backed environment)' -ForegroundColor DarkGray
}
Write-Host "  Output folder : $OutputFolder" -ForegroundColor White
Write-Host ''

###############################################################################
# 4.  DATA STRUCTURES
###############################################################################
$detailRows     = [System.Collections.Generic.List[PSObject]]::new()
$tenantShares   = [System.Collections.Generic.List[PSObject]]::new()
$groupMembers   = [System.Collections.Generic.List[PSObject]]::new()
$unresolvedList = [System.Collections.Generic.List[PSObject]]::new()

# GroupId  --> DisplayName
$groupsMap = [System.Collections.Generic.Dictionary[string,string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)

# Entra ObjectId --> cached principal info  (avoids repeated Graph calls)
$principalCache = [System.Collections.Generic.Dictionary[string,PSObject]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)

###############################################################################
# 5.  HELPER FUNCTIONS
###############################################################################

function Add-DetailRow {
    param(
        [string]$ArtifactType,
        [string]$ArtifactName,
        [string]$ArtifactId,
        [string]$AccessType,
        [string]$PrincipalType,
        [string]$PrincipalId,
        [string]$PrincipalDisplayName,
        [string]$PrincipalUPN,
        [string]$Source
    )
    $null = $script:detailRows.Add([PSCustomObject][ordered]@{
        EnvironmentId          = $EnvironmentId
        EnvironmentDisplayName = $envDisplayName
        ArtifactType           = $ArtifactType
        ArtifactName           = $ArtifactName
        ArtifactId             = $ArtifactId
        AccessType             = $AccessType
        PrincipalType          = $PrincipalType
        PrincipalId            = $PrincipalId
        PrincipalDisplayName   = $PrincipalDisplayName
        PrincipalUPN           = $PrincipalUPN
        Source                 = $Source
    })
}

function Add-TenantShare {
    param(
        [string]$ArtifactType,
        [string]$ArtifactName,
        [string]$ArtifactId,
        [string]$ShareType
    )
    $null = $script:tenantShares.Add([PSCustomObject][ordered]@{
        EnvironmentId          = $EnvironmentId
        EnvironmentDisplayName = $envDisplayName
        ArtifactType           = $ArtifactType
        ArtifactName           = $ArtifactName
        ArtifactId             = $ArtifactId
        ShareType              = $ShareType
    })
}

function Register-Group {
    param([string]$GroupId, [string]$GroupDisplayName, [string]$Source)
    if (-not [string]::IsNullOrWhiteSpace($GroupId) -and
        -not $script:groupsMap.ContainsKey($GroupId)) {
        $script:groupsMap[$GroupId] = if ($GroupDisplayName) { $GroupDisplayName } else { $GroupId }
    }
}

function Add-Unresolved {
    param(
        [string]$PrincipalId,
        [string]$PrincipalType,
        [string]$ArtifactName,
        [string]$Source
    )
    $null = $script:unresolvedList.Add([PSCustomObject][ordered]@{
        PrincipalId   = $PrincipalId
        PrincipalType = $PrincipalType
        ArtifactName  = $ArtifactName
        Source        = $Source
    })
}

# Returns $true when the Graph error is a 404 / ResourceNotFound. That is expected
# when probing an object as the wrong type; any other status must surface.
function Test-IsNotFoundError {
    param([System.Management.Automation.ErrorRecord]$Err)
    $msg = $Err.Exception.Message
    if ($msg -match 'Request_ResourceNotFound|ResourceNotFound|does not exist|404') {
        return $true
    }
    # Graph SDK wraps the HTTP response inside ODataError; check status code too
    $inner = $Err.Exception.InnerException
    while ($inner) {
        if ($inner -is [System.Net.WebException]) {
            $response = $inner.Response
            if ($response -is [System.Net.HttpWebResponse]) {
                if ([int]$response.StatusCode -eq 404) { return $true }
            }
        }
        $inner = $inner.InnerException
    }
    return $false
}

# Resolve an Entra object ID to display-name / UPN / type via Microsoft Graph.
# Returns $null if the object cannot be resolved.
function Resolve-Principal {
    param([string]$ObjectId)
    if ([string]::IsNullOrWhiteSpace($ObjectId)) { return $null }
    if ($script:principalCache.ContainsKey($ObjectId)) {
        return $script:principalCache[$ObjectId]
    }

    $r = $null

    # Try user first
    try {
        $u = Get-MgUser -UserId $ObjectId `
            -Property 'id,displayName,userPrincipalName' -ErrorAction Stop
        $r = [PSCustomObject]@{
            ObjectId    = $ObjectId
            DisplayName = $u.DisplayName
            UPN         = $u.UserPrincipalName
            Type        = 'User'
        }
    } catch {
        # Only suppress "not found"; the object may be a group or service principal.
        # Anything else, such as auth failure, throttling, or network errors, must be visible.
        if (-not (Test-IsNotFoundError -Err $_)) {
            Write-Warning "  Resolve-Principal: Get-MgUser failed for '$ObjectId': $($_.Exception.Message)"
        }
    }

    # Try group
    if (-not $r) {
        try {
            $g = Get-MgGroup -GroupId $ObjectId `
                -Property 'id,displayName,mail' -ErrorAction Stop
            $r = [PSCustomObject]@{
                ObjectId    = $ObjectId
                DisplayName = $g.DisplayName
                UPN         = $g.Mail
                Type        = 'Group'
            }
        } catch {
            if (-not (Test-IsNotFoundError -Err $_)) {
                Write-Warning "  Resolve-Principal: Get-MgGroup failed for '$ObjectId': $($_.Exception.Message)"
            }
        }
    }

    if ($r) { $script:principalCache[$ObjectId] = $r }
    return $r
}

# GET all pages from a Dataverse OData endpoint, following @odata.nextLink.
# Returns a plain PowerShell array.
function Invoke-DvGet {
    param([string]$Uri, [hashtable]$Headers)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    do {
        try {
            $resp = Invoke-RestMethod -Uri $next -Headers $Headers `
                -Method Get -ErrorAction Stop
            if ($null -ne $resp.value) {
                foreach ($v in $resp.value) { $null = $items.Add($v) }
            }
            $next = $resp.'@odata.nextLink'
        } catch {
            Write-Warning "    Dataverse GET failed: $($_.Exception.Message)"
            Write-Warning "    URI: $next"
            break
        }
    } while (-not [string]::IsNullOrEmpty($next))
    return @($items)
}

# Get transitive user members of an Entra group via Microsoft Graph.
# Returns a plain PowerShell array of member objects.
function Expand-Group {
    param([string]$GroupId, [string]$GroupDisplayName)
    $out = [System.Collections.Generic.List[PSObject]]::new()
    try {
        $raw = Get-MgGroupTransitiveMember -GroupId $GroupId -All -ErrorAction Stop
        foreach ($m in $raw) {
            $oType = $m.AdditionalProperties['@odata.type']
            if ($oType -ne '#microsoft.graph.user') { continue }
            $upn = $m.AdditionalProperties['userPrincipalName']
            $dn  = $m.AdditionalProperties['displayName']
            $null = $out.Add([PSCustomObject][ordered]@{
                GroupId           = $GroupId
                GroupDisplayName  = $GroupDisplayName
                MemberId          = $m.Id
                MemberDisplayName = $dn
                MemberUPN         = $upn
            })
            # Cache for use in seed-file building
            if (-not $script:principalCache.ContainsKey($m.Id)) {
                $script:principalCache[$m.Id] = [PSCustomObject]@{
                    ObjectId    = $m.Id
                    DisplayName = $dn
                    UPN         = $upn
                    Type        = 'User'
                }
            }
        }
    } catch {
        Write-Warning "    Could not expand group '$GroupDisplayName' ($GroupId): $($_.Exception.Message)"
    }
    return @($out)
}

# Acquire a Dataverse bearer token via MSAL interactive browser sign-in.
# Microsoft.Identity.Client ships with the Microsoft.Graph module - no extra dependencies.
function Get-DvToken {
    param([Parameter(Mandatory)][string]$InstanceUrl, [string]$TenantHint = 'organizations')
    $scopes    = [string[]]@("$($InstanceUrl.TrimEnd('/'))/.default")
    $clientId  = '1950a258-227b-4e31-a9cf-717495945fc2'   # Microsoft Azure PowerShell public client
    $authority = "https://login.microsoftonline.com/$TenantHint"

    $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($clientId)
    $builder = $builder.WithAuthority($authority)
    $builder = $builder.WithDefaultRedirectUri()
    $app     = $builder.Build()

    # Try silent first (cache hit from earlier interactive sign-in)
    $accounts = $app.GetAccountsAsync().GetAwaiter().GetResult()
    if ($accounts) {
        try {
            $acqSilent = $app.AcquireTokenSilent($scopes, ($accounts | Select-Object -First 1))
            $result    = $acqSilent.ExecuteAsync().GetAwaiter().GetResult()
            return [string]$result.AccessToken
        } catch { }
    }

    # Interactive browser pop-up
    Write-Host ''
    Write-Host '  Dataverse sign-in: a browser window will open for authentication.' -ForegroundColor Yellow
    $acqInteractive = $app.AcquireTokenInteractive($scopes)
    $result         = $acqInteractive.ExecuteAsync().GetAwaiter().GetResult()
    return [string]$result.AccessToken
}

###############################################################################
# 6.  CANVAS APPS
###############################################################################
Write-Host '  -- Canvas Apps --------------------------------------------------------' -ForegroundColor Yellow

$apps = @(Get-AdminPowerApp -EnvironmentName $EnvironmentId)
Write-Host "  Found $($apps.Count) canvas app(s)." -ForegroundColor White

for ($i = 0; $i -lt $apps.Count; $i++) {
    $app = $apps[$i]
    $pct = [int](($i + 1) / [Math]::Max(1, $apps.Count) * 100)
    Write-Progress -Activity 'Canvas Apps' -Status $app.DisplayName -PercentComplete $pct

    try {
        $roles = @(Get-AdminPowerAppRoleAssignment `
            -EnvironmentName $EnvironmentId -AppName $app.AppName -ErrorAction Stop)
    } catch {
        Write-Warning "  Cannot retrieve roles for app '$($app.DisplayName)': $($_.Exception.Message)"
        continue
    }

    if (-not $roles) { continue }
    foreach ($role in $roles) {
        $pType = if ($role.PrincipalType) { [string]$role.PrincipalType } else { 'Unknown' }

        $pObjId  = if ($role.PrincipalObjectId) {
                    $role.PrincipalObjectId                
                } else { '' }

        $pDn  = ''

        $pUpn = if ($role.PrincipalEmail) {
                    $role.PrincipalEmail                
                } else { '' }

        if ($pType -eq 'Tenant') {
            Add-TenantShare -ArtifactType 'CanvasApp' `
                -ArtifactName $app.DisplayName `
                -ArtifactId   $app.AppName `
                -ShareType    'EntireTenant'
            Add-DetailRow -ArtifactType 'CanvasApp' `
                -ArtifactName $app.DisplayName `
                -ArtifactId   $app.AppName `
                -AccessType   $role.RoleName `
                -PrincipalType 'Tenant' -PrincipalId '' `
                -PrincipalDisplayName 'Entire Tenant' -PrincipalUPN '' `
                -Source 'CanvasApp'
            continue
        }

        if ($pType -eq 'Group' -and -not [string]::IsNullOrWhiteSpace($pObjId)) {
            Register-Group -GroupId $pObjId -GroupDisplayName $pDn -Source 'CanvasApp'
        }

        if ([string]::IsNullOrWhiteSpace($pDn) -and
            -not [string]::IsNullOrWhiteSpace($pObjId)) {
            $resolved = Resolve-Principal -ObjectId $pObjId
            if ($resolved) {
                $pDn  = $resolved.DisplayName
                $pUpn = $resolved.UPN
            } else {
                Add-Unresolved -PrincipalId $pObjId -PrincipalType $pType `
                    -ArtifactName $app.DisplayName -Source 'CanvasApp'
            }
        } elseif ([string]::IsNullOrWhiteSpace($pUpn) -and
                  -not [string]::IsNullOrWhiteSpace($pObjId)) {
            $resolved = Resolve-Principal -ObjectId $pObjId
            if ($resolved) { $pUpn = $resolved.UPN }
        }

        Add-DetailRow -ArtifactType 'CanvasApp' `
            -ArtifactName $app.DisplayName `
            -ArtifactId   $app.AppName `
            -AccessType   $role.RoleName `
            -PrincipalType $pType -PrincipalId $pObjId `
            -PrincipalDisplayName $pDn -PrincipalUPN $pUpn `
            -Source 'CanvasApp'
    }
}
Write-Progress -Activity 'Canvas Apps' -Completed
Write-Host "  Done. ($($detailRows.Count) detail rows so far)" -ForegroundColor Green
Write-Host ''

###############################################################################
# 7.  CLOUD FLOWS
###############################################################################
Write-Host '  -- Cloud Flows --------------------------------------------------------' -ForegroundColor Yellow

$flows = @(Get-AdminFlow -EnvironmentName $EnvironmentId)
Write-Host "  Found $($flows.Count) cloud flow(s)." -ForegroundColor White

for ($i = 0; $i -lt $flows.Count; $i++) {
    $flow = $flows[$i]
    $pct  = [int](($i + 1) / [Math]::Max(1, $flows.Count) * 100)
    Write-Progress -Activity 'Cloud Flows' -Status $flow.DisplayName -PercentComplete $pct

    try {
        $ownerRoles = @(Get-AdminFlowOwnerRole `
            -EnvironmentName $EnvironmentId -FlowName $flow.FlowName -ErrorAction Stop)
    } catch {
        Write-Warning "  Cannot retrieve owner roles for flow '$($flow.DisplayName)': $($_.Exception.Message)"
        continue
    }

    if (-not $ownerRoles) { continue }
    foreach ($role in $ownerRoles) {
        # $role.PrincipalType, $role.PrincipalObjectId, $role.PrincipalDisplayName
        $pType = if ($role.PrincipalType) { [string]$role.PrincipalType } else { 'Unknown' }
        $pObjId   = if ($role.PrincipalObjectId) { $role.PrincipalObjectId } else { '' }
        $pDn   = ''
        $pUpn  = ''

        if ($pType -eq 'Group' -and -not [string]::IsNullOrWhiteSpace($pObjId)) {
            Register-Group -GroupId $pObjId -GroupDisplayName $pDn -Source 'CloudFlow'
        }

        if ([string]::IsNullOrWhiteSpace($pDn) -and
            -not [string]::IsNullOrWhiteSpace($pObjId)) {
            $resolved = Resolve-Principal -ObjectId $pObjId
            if ($resolved) {
                $pDn  = $resolved.DisplayName
                $pUpn = $resolved.UPN
            } else {
                Add-Unresolved -PrincipalId $pObjId -PrincipalType $pType `
                    -ArtifactName $flow.DisplayName -Source 'CloudFlow'
            }
        } elseif (-not [string]::IsNullOrWhiteSpace($pObjId)) {
            $resolved = Resolve-Principal -ObjectId $pObjId
            if ($resolved) { $pUpn = $resolved.UPN }
        }

        Add-DetailRow -ArtifactType 'CloudFlow' `
            -ArtifactName $flow.DisplayName `
            -ArtifactId   $flow.FlowName `
            -AccessType   $role.RoleName `
            -PrincipalType $pType -PrincipalId $pObjId `
            -PrincipalDisplayName $pDn -PrincipalUPN $pUpn `
            -Source 'CloudFlow'
    }
}
Write-Progress -Activity 'Cloud Flows' -Completed
Write-Host "  Done. ($($detailRows.Count) detail rows so far)" -ForegroundColor Green
Write-Host ''

###############################################################################
# 8.  CONNECTIONS
###############################################################################
Write-Host '  -- Connections --------------------------------------------------------' -ForegroundColor Yellow

try {
    $connections = @(Get-AdminPowerAppConnection -EnvironmentName $EnvironmentId -ErrorAction Stop)
} catch {
    Write-Warning "  Get-AdminPowerAppConnection failed: $($_.Exception.Message). Skipping."
    $connections = @()
}
Write-Host "  Found $($connections.Count) connection(s)." -ForegroundColor White

foreach ($conn in $connections) {
    $cb = if ($conn.PSObject.Properties['CreatedBy']) { $conn.CreatedBy } else { $null }
    if (-not $cb) { continue }

    $pObjId  = if ($cb.PSObject.Properties['id'])          { $cb.id }          else { '' }
    $pDn  = if ($cb.PSObject.Properties['displayName']) { $cb.displayName } else { '' }
    $pUpn = if ($cb.PSObject.Properties['email'])       { $cb.email }       else { '' }

    if ([string]::IsNullOrWhiteSpace($pDn) -and
        -not [string]::IsNullOrWhiteSpace($pObjId)) {
        $resolved = Resolve-Principal -ObjectId $pObjId
        if ($resolved) {
            $pDn  = $resolved.DisplayName
            $pUpn = $resolved.UPN
        } else {
            Add-Unresolved -PrincipalId $pObjId -PrincipalType 'User' `
                -ArtifactName "$($conn.DisplayName) [$($conn.ConnectionName)]" `
                -Source 'Connection'
        }
    } elseif ([string]::IsNullOrWhiteSpace($pUpn) -and
              -not [string]::IsNullOrWhiteSpace($pObjId)) {
        $resolved = Resolve-Principal -ObjectId $pObjId
        if ($resolved) { $pUpn = $resolved.UPN }
    }

    Add-DetailRow -ArtifactType 'Connection' `
        -ArtifactName "$($conn.DisplayName) [$($conn.ConnectionName)]" `
        -ArtifactId   $conn.ConnectionName `
        -AccessType   'Owner' `
        -PrincipalType 'User' -PrincipalId $pObjId `
        -PrincipalDisplayName $pDn -PrincipalUPN $pUpn `
        -Source 'Connection'
}
Write-Host "  Done. ($($detailRows.Count) detail rows so far)" -ForegroundColor Green
Write-Host ''

###############################################################################
# 9.  CUSTOM CONNECTORS
###############################################################################
Write-Host '  -- Custom Connectors --------------------------------------------------' -ForegroundColor Yellow

$connectors = @()
try {
    $connectors = @(Get-AdminPowerAppConnector -EnvironmentName $EnvironmentId -ErrorAction Stop |
        Where-Object { $null -ne $_.Internal -and $_.Internal.isCustomApi -eq $true })
} catch {
    Write-Warning "  Get-AdminPowerAppConnector failed: $($_.Exception.Message). Skipping."
}
Write-Host "  Found $($connectors.Count) custom connector(s)." -ForegroundColor White

foreach ($c in $connectors) {
    $cb = if ($c.PSObject.Properties['CreatedBy']) { $c.CreatedBy } else { $null }
    if (-not $cb) { continue }

    $pObjId  = if ($cb.PSObject.Properties['id'])          { $cb.id }          else { '' }
    $pDn  = if ($cb.PSObject.Properties['displayName']) { $cb.displayName } else { '' }
    $pUpn = if ($cb.PSObject.Properties['email'])       { $cb.email }       else { '' }

    if ([string]::IsNullOrWhiteSpace($pDn) -and
        -not [string]::IsNullOrWhiteSpace($pObjId)) {
        $resolved = Resolve-Principal -ObjectId $pObjId
        if ($resolved) {
            $pDn  = $resolved.DisplayName
            $pUpn = $resolved.UPN
        } else {
            Add-Unresolved -PrincipalId $pObjId -PrincipalType 'User' `
                -ArtifactName $c.DisplayName -Source 'CustomConnector'
        }
    } elseif ([string]::IsNullOrWhiteSpace($pUpn) -and
              -not [string]::IsNullOrWhiteSpace($pObjId)) {
        $resolved = Resolve-Principal -ObjectId $pObjId
        if ($resolved) { $pUpn = $resolved.UPN }
    }

    Add-DetailRow -ArtifactType 'CustomConnector' `
        -ArtifactName $c.DisplayName `
        -ArtifactId   $c.ConnectorName `
        -AccessType   'Owner' `
        -PrincipalType 'User' -PrincipalId $pObjId `
        -PrincipalDisplayName $pDn -PrincipalUPN $pUpn `
        -Source 'CustomConnector'
}
Write-Host "  Done. ($($detailRows.Count) detail rows so far)" -ForegroundColor Green
Write-Host ''

###############################################################################
# 10.  DATAVERSE  (conditional on -IncludeDataverse)
###############################################################################
if ($IncludeDataverse) {
    if ([string]::IsNullOrWhiteSpace($instanceUrl)) {
        Write-Warning '-IncludeDataverse specified but environment has no Dataverse instance. Skipping Dataverse queries.'
    } else {
        Write-Host '  -- Dataverse ----------------------------------------------------------' -ForegroundColor Yellow
        $dvBase = $instanceUrl.TrimEnd('/')

        # -- Obtain Dataverse bearer token -----------------------------------
        Write-Host '  Obtaining Dataverse bearer token ...' -NoNewline -ForegroundColor White
        $dvToken = $null
        try {
            $tenantHint = if ($TenantId) { $TenantId } else { 'organizations' }
            $dvToken = Get-DvToken -InstanceUrl $dvBase -TenantHint $tenantHint
            Write-Host '  Dataverse token acquired.' -ForegroundColor Green
        } catch {
            Write-Host ' FAILED' -ForegroundColor Red
            Write-Warning "  Cannot obtain Dataverse token: $($_.Exception.Message)"
            Write-Warning '  Skipping Dataverse queries.'
        }

        if ($dvToken) {
            $dvH = @{
                Authorization      = "Bearer $dvToken"
                'OData-MaxVersion' = '4.0'
                'OData-Version'    = '4.0'
                Accept             = 'application/json'
                Prefer             = 'odata.maxpagesize=5000'
            }

            # -- 10a. Users with at least one security role ------------------
            Write-Host '  [DV 1/4] Users with at least one security role ...' -ForegroundColor White
            $usersUrl = $dvBase +
                '/api/data/v9.2/systemusers' +
                '?$select=fullname,domainname,internalemailaddress,systemuserid,azureactivedirectoryobjectid,isdisabled,accessmode' +
                '&$filter=isdisabled eq false and accessmode ne 3' +
                '&$expand=systemuserroles_association($select=name,roleid)'

            $dvUsers         = Invoke-DvGet -Uri $usersUrl -Headers $dvH
            $dvUsersWithRoles = @($dvUsers | Where-Object {
                $_.systemuserroles_association -and
                $_.systemuserroles_association.Count -gt 0
            })
            Write-Host "    $($dvUsersWithRoles.Count) user(s) hold at least one security role." -ForegroundColor White

            foreach ($u in $dvUsersWithRoles) {
                $roleNames = ($u.systemuserroles_association |
                              Select-Object -ExpandProperty name) -join '; '
                $aadId     = $u.azureactivedirectoryobjectid
                $upn       = $u.domainname          # Dataverse: domainname == Entra UPN
                $dn        = $u.fullname

                if ($aadId -and -not $principalCache.ContainsKey($aadId)) {
                    $principalCache[$aadId] = [PSCustomObject]@{
                        ObjectId = $aadId; DisplayName = $dn; UPN = $upn; Type = 'User'
                    }
                }

                Add-DetailRow -ArtifactType 'DataverseSecurityRole' `
                    -ArtifactName $roleNames -ArtifactId $u.systemuserid `
                    -AccessType   'SecurityRoleHolder' -PrincipalType 'User' `
                    -PrincipalId  $aadId -PrincipalDisplayName $dn `
                    -PrincipalUPN $upn -Source 'Dataverse'
            }

            # -- 10b. Teams with security roles ------------------------------
            Write-Host '  [DV 2/4] Teams with security roles ...' -ForegroundColor White
            $teamsUrl = $dvBase +
                '/api/data/v9.2/teams' +
                '?$select=name,teamtype,azureactivedirectoryobjectid,teamid' +
                '&$expand=teamroles_association($select=name,roleid)'

            $dvTeams         = Invoke-DvGet -Uri $teamsUrl -Headers $dvH
            $dvTeamsWithRoles = @($dvTeams | Where-Object {
                $_.teamroles_association -and $_.teamroles_association.Count -gt 0
            })
            Write-Host "    $($dvTeamsWithRoles.Count) team(s) hold at least one security role." -ForegroundColor White

            foreach ($team in $dvTeamsWithRoles) {
                $roleNames = ($team.teamroles_association |
                              Select-Object -ExpandProperty name) -join '; '
                # teamtype: 0=Owner, 1=Access, 2=AAD (Entra group), 3=AAD Office group
                $teamType  = $team.teamtype
                $aadObjId  = $team.azureactivedirectoryobjectid

                if ($teamType -in @(2, 3)) {
                    # Entra-group-backed team - report as a group
                    if ($aadObjId) {
                        Register-Group -GroupId $aadObjId `
                            -GroupDisplayName $team.name -Source 'DataverseTeam'
                    }
                    Add-DetailRow -ArtifactType 'DataverseSecurityRole' `
                        -ArtifactName $roleNames -ArtifactId $team.teamid `
                        -AccessType   'TeamSecurityRoleHolder' `
                        -PrincipalType 'Group' -PrincipalId $aadObjId `
                        -PrincipalDisplayName $team.name -PrincipalUPN '' `
                        -Source 'DataverseTeam'
                } else {
                    # Owner/Access team - enumerate individual members
                    $membUrl = "$dvBase/api/data/v9.2/teams($($team.teamid))" +
                               '/teammembership_association' +
                               '?$select=fullname,domainname,azureactivedirectoryobjectid,systemuserid'
                    $teamMembers = Invoke-DvGet -Uri $membUrl -Headers $dvH

                    if (-not $teamMembers) { continue }
                    foreach ($tm in $teamMembers) {
                        $aadId = $tm.azureactivedirectoryobjectid
                        $upn   = $tm.domainname
                        $dn    = $tm.fullname

                        if ($aadId -and -not $principalCache.ContainsKey($aadId)) {
                            $principalCache[$aadId] = [PSCustomObject]@{
                                ObjectId = $aadId; DisplayName = $dn; UPN = $upn; Type = 'User'
                            }
                        }

                        Add-DetailRow -ArtifactType 'DataverseSecurityRole' `
                            -ArtifactName "$($team.name) -> $roleNames" `
                            -ArtifactId   $team.teamid `
                            -AccessType   'TeamMember' -PrincipalType 'User' `
                            -PrincipalId  $aadId -PrincipalDisplayName $dn `
                            -PrincipalUPN $upn -Source 'DataverseTeam'
                    }
                }
            }

            # -- 10c. Copilot Studio agents ----------------------------------
            Write-Host '  [DV 3/4] Copilot Studio agents ...' -ForegroundColor White
            $botsUrl = $dvBase +
                '/api/data/v9.2/bots' +
                '?$select=name,botid,_ownerid_value' +
                '&$expand=owninguser($select=fullname,domainname,azureactivedirectoryobjectid)'
            try {
                $dvBots = Invoke-DvGet -Uri $botsUrl -Headers $dvH
                Write-Host "    $($dvBots.Count) Copilot Studio agent(s) found." -ForegroundColor White

                foreach ($bot in $dvBots) {
                    # -- Owner ----------------------------------------------
                    $owner = $bot.owninguser
                    if ($owner) {
                        $aadId = $owner.azureactivedirectoryobjectid
                        $upn   = $owner.domainname
                        $dn    = $owner.fullname

                        if ($aadId -and -not $principalCache.ContainsKey($aadId)) {
                            $principalCache[$aadId] = [PSCustomObject]@{
                                ObjectId = $aadId; DisplayName = $dn; UPN = $upn; Type = 'User'
                            }
                        }
                        Add-DetailRow -ArtifactType 'CopilotStudioAgent' `
                            -ArtifactName $bot.name -ArtifactId $bot.botid `
                            -AccessType   'Owner' -PrincipalType 'User' `
                            -PrincipalId  $aadId -PrincipalDisplayName $dn `
                            -PrincipalUPN $upn -Source 'CopilotStudioAgent'
                    }

                    # -- Shared-with principals -----------------------------
                    # Uses the standard Dataverse RetrieveSharedPrincipalsAndAccess action
                    try {
                        $ref      = [Uri]::EscapeDataString(
                                        "{'@odata.id':'bots($($bot.botid))'}")
                        $shareUrl = "$dvBase/api/data/v9.2/RetrieveSharedPrincipalsAndAccess" +
                                    "(Target=@t)?@t=$ref"
                        $shareResp = Invoke-RestMethod -Uri $shareUrl `
                            -Headers $dvH -Method Get -ErrorAction Stop

                        if ($shareResp.PrincipalAccesses) {
                            foreach ($pa in $shareResp.PrincipalAccesses) {
                                $pr     = $pa.Principal
                                $prOdt  = $pr.'@odata.type'
                                $prType = if ($prOdt -match 'systemuser') { 'User'  }
                                          elseif ($prOdt -match 'team')   { 'Group' }
                                          else                             { 'Unknown' }

                                $prAadId = if ($pr.PSObject.Properties[
                                               'azureactivedirectoryobjectid']) {
                                               $pr.azureactivedirectoryobjectid } else { '' }
                                $prDn    = if ($pr.PSObject.Properties['fullname'] -and
                                               $pr.fullname) {
                                               $pr.fullname }
                                           elseif ($pr.PSObject.Properties['name']) {
                                               $pr.name } else { '' }
                                $prUpn   = if ($pr.PSObject.Properties['domainname']) {
                                               $pr.domainname } else { '' }

                                if ($prType -eq 'Group' -and $prAadId) {
                                    Register-Group -GroupId $prAadId `
                                        -GroupDisplayName $prDn -Source 'CopilotStudioAgent'
                                }

                                Add-DetailRow -ArtifactType 'CopilotStudioAgent' `
                                    -ArtifactName $bot.name -ArtifactId $bot.botid `
                                    -AccessType   'SharedWith' -PrincipalType $prType `
                                    -PrincipalId  $prAadId -PrincipalDisplayName $prDn `
                                    -PrincipalUPN $prUpn -Source 'CopilotStudioAgent'
                            }
                        }
                    } catch {
                        Write-Verbose "    RetrieveSharedPrincipalsAndAccess failed for agent '$($bot.name)': $($_.Exception.Message)"
                    }
                }
            } catch {
                Write-Warning "  Could not query Copilot Studio agents (bots table): $($_.Exception.Message)"
            }

            # -- 10d. Solution cloud-flow owners (workflow category 5) --------
            Write-Host '  [DV 4/4] Solution cloud-flow owners (workflow category 5) ...' -ForegroundColor White
            $wfUrl = $dvBase +
                '/api/data/v9.2/workflows' +
                '?$select=name,workflowid,_ownerid_value' +
                '&$filter=category eq 5' +
                '&$expand=owninguser($select=fullname,domainname,azureactivedirectoryobjectid)'
            try {
                $dvFlows = Invoke-DvGet -Uri $wfUrl -Headers $dvH
                Write-Host "    $($dvFlows.Count) solution cloud flow(s) found." -ForegroundColor White

                foreach ($wf in $dvFlows) {
                    $owner = $wf.owninguser
                    if ($owner) {
                        $aadId = $owner.azureactivedirectoryobjectid
                        $upn   = $owner.domainname
                        $dn    = $owner.fullname

                        if ($aadId -and -not $principalCache.ContainsKey($aadId)) {
                            $principalCache[$aadId] = [PSCustomObject]@{
                                ObjectId = $aadId; DisplayName = $dn; UPN = $upn; Type = 'User'
                            }
                        }
                        Add-DetailRow -ArtifactType 'SolutionCloudFlow' `
                            -ArtifactName $wf.name -ArtifactId $wf.workflowid `
                            -AccessType   'Owner' -PrincipalType 'User' `
                            -PrincipalId  $aadId -PrincipalDisplayName $dn `
                            -PrincipalUPN $upn -Source 'DataverseSolutionFlow'
                    } else {
                        # Owner may be a team - log raw _ownerid_value as unresolved
                        $ownerRaw = $wf.'_ownerid_value'
                        if ($ownerRaw) {
                            Add-Unresolved -PrincipalId $ownerRaw `
                                -PrincipalType 'Team/Unknown' `
                                -ArtifactName $wf.name -Source 'DataverseSolutionFlow'
                        }
                    }
                }
            } catch {
                Write-Warning "  Could not query solution cloud flows: $($_.Exception.Message)"
            }

            Write-Host "  Dataverse done. ($($detailRows.Count) detail rows so far)" -ForegroundColor Green
            Write-Host ''
        }
    }
}

###############################################################################
# 11.  GROUP EXPANSION  (conditional on -ExpandGroups)
###############################################################################
if ($ExpandGroups) {
    Write-Host '  -- Group Expansion ----------------------------------------------------' -ForegroundColor Yellow
    if ($groupsMap.Count -eq 0) {
        Write-Host '  No groups encountered; nothing to expand.' -ForegroundColor White
    } else {
        Write-Host "  Expanding $($groupsMap.Count) group(s) ..." -ForegroundColor White
        $gi   = 0
        $keys = @($groupsMap.Keys)
        foreach ($gId in $keys) {
            $gi++
            $gDn = $groupsMap[$gId]
            $pct = [int]($gi / [Math]::Max(1, $keys.Count) * 100)
            Write-Progress -Activity 'Expanding groups' -Status $gDn -PercentComplete $pct
            $members = Expand-Group -GroupId $gId -GroupDisplayName $gDn
            foreach ($m in $members) { $null = $groupMembers.Add($m) }
        }
        Write-Progress -Activity 'Expanding groups' -Completed
        Write-Host "  Done. $($groupMembers.Count) transitive user membership(s) found." -ForegroundColor Green
    }
    Write-Host ''
}

###############################################################################
# 12.  BUILD SUMMARY  (one row per distinct principal)
###############################################################################
Write-Host '  -- Building summary ---------------------------------------------------' -ForegroundColor Yellow

$summaryMap = [System.Collections.Generic.Dictionary[string,PSObject]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)
$summarySourceMap = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.HashSet[string]]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)

function ConvertTo-SourceDetailsArray {
    param([System.Collections.Generic.HashSet[string]]$Values)
    $items = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)
    if ($items.Count -eq 0) { return '[]' }
    return "[$($items -join ',')]"
}

function ConvertTo-SourceDetailPair {
    param([string]$Type, [string]$Name, [string]$Id)
    $safeType = if ($Type) { $Type.Replace('"', '\"') } else { 'Unknown' }
    $safeName = if ($Name) { $Name.Replace('"', '\"') } else { '' }
    $safeId = if ($Id) { $Id.Replace('"', '\"') } else { '' }
    $value = if ([string]::IsNullOrWhiteSpace($safeId)) { $safeName } else { "$safeName - $safeId" }
    return "`"$safeType`":`"$value`""
}

foreach ($row in $detailRows) {
    if ($row.PrincipalType -eq 'Tenant' -or
        [string]::IsNullOrWhiteSpace($row.PrincipalId)) { continue }

    $key = "$($row.PrincipalType)|$($row.PrincipalId)"
    $sourceType = if ($row.Source -eq $row.ArtifactType) { $row.Source } else { $row.ArtifactType }
    $sourceDetail = ConvertTo-SourceDetailPair -Type $sourceType -Name $row.ArtifactName -Id $row.ArtifactId

    if (-not $summarySourceMap.ContainsKey($key)) {
        $summarySourceMap[$key] = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase)
    }
    $null = $summarySourceMap[$key].Add($sourceDetail)

    if ($summaryMap.ContainsKey($key)) {
        $existing = $summaryMap[$key]
        $existing.SourceDetails = ConvertTo-SourceDetailsArray -Values $summarySourceMap[$key]
        # Backfill better display-name / UPN if we have it now
        if ([string]::IsNullOrWhiteSpace($existing.PrincipalDisplayName) -and
            -not [string]::IsNullOrWhiteSpace($row.PrincipalDisplayName)) {
            $existing.PrincipalDisplayName = $row.PrincipalDisplayName
        }
        if ([string]::IsNullOrWhiteSpace($existing.PrincipalUPN) -and
            -not [string]::IsNullOrWhiteSpace($row.PrincipalUPN)) {
            $existing.PrincipalUPN = $row.PrincipalUPN
        }
    } else {
        $summaryMap[$key] = [PSCustomObject][ordered]@{
            PrincipalId          = $row.PrincipalId
            PrincipalDisplayName = $row.PrincipalDisplayName
            PrincipalUPN         = $row.PrincipalUPN
            PrincipalType        = $row.PrincipalType
            SourceDetails        = ConvertTo-SourceDetailsArray -Values $summarySourceMap[$key]
        }
    }
}

$summaryRows = @($summaryMap.Values | Sort-Object PrincipalDisplayName)
Write-Host "  $($summaryRows.Count) distinct principal(s)." -ForegroundColor Green
Write-Host ''

###############################################################################
# 13.  BUILD SECURITY GROUP SEED  (Entra portal bulk-import format)
###############################################################################
Write-Host '  -- Building security-group seed file ---------------------------------' -ForegroundColor Yellow

$seedUpns = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)

# From detail rows - user principals with a valid UPN
foreach ($row in $detailRows) {
    if ($row.PrincipalType -eq 'User' -and
        -not [string]::IsNullOrWhiteSpace($row.PrincipalUPN) -and
        $row.PrincipalUPN -match '@') {
        $null = $seedUpns.Add($row.PrincipalUPN.Trim())
    }
}

# From expanded group members
foreach ($gm in $groupMembers) {
    if (-not [string]::IsNullOrWhiteSpace($gm.MemberUPN) -and
        $gm.MemberUPN -match '@') {
        $null = $seedUpns.Add($gm.MemberUPN.Trim())
    }
}

$sortedSeedUpns = @($seedUpns | Sort-Object)
Write-Host "  $($sortedSeedUpns.Count) unique UPN(s) for seed file." -ForegroundColor Green
Write-Host ''

###############################################################################
# 14.  EXPORT CSVs
###############################################################################
Write-Host '  -- Exporting CSVs -----------------------------------------------------' -ForegroundColor Yellow

# UserAccess_Detail.csv
$p = Join-Path $OutputFolder 'UserAccess_Detail.csv'
$detailRows | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
Write-Host "  UserAccess_Detail.csv         ($($detailRows.Count) rows)" -ForegroundColor White

# UserAccess_Summary.csv
$p = Join-Path $OutputFolder 'UserAccess_Summary.csv'
if ($summaryRows.Count -gt 0) {
    $summaryRows | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
} else {
    [PSCustomObject][ordered]@{
        PrincipalId=''; PrincipalDisplayName=''; PrincipalUPN=''; PrincipalType=''; SourceDetails='[]'
    } | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
}
Write-Host "  UserAccess_Summary.csv        ($($summaryRows.Count) rows)" -ForegroundColor White

# TenantWideShares.csv
$p = Join-Path $OutputFolder 'TenantWideShares.csv'
if ($tenantShares.Count -gt 0) {
    $tenantShares | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
} else {
    [PSCustomObject][ordered]@{
        EnvironmentId=''; EnvironmentDisplayName=''; ArtifactType=''
        ArtifactName=''; ArtifactId=''; ShareType=''
    } | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
}
Write-Host "  TenantWideShares.csv          ($($tenantShares.Count) rows)" -ForegroundColor White

# GroupsEncountered.csv
$p = Join-Path $OutputFolder 'GroupsEncountered.csv'
$groupRows = @($groupsMap.GetEnumerator() | ForEach-Object {
    [PSCustomObject][ordered]@{ GroupId = $_.Key; GroupDisplayName = $_.Value }
})
if ($groupRows.Count -gt 0) {
    $groupRows | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
} else {
    [PSCustomObject][ordered]@{ GroupId = ''; GroupDisplayName = '' } |
        Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
}
Write-Host "  GroupsEncountered.csv         ($($groupsMap.Count) group(s))" -ForegroundColor White

# GroupMembers.csv
$p = Join-Path $OutputFolder 'GroupMembers.csv'
if ($groupMembers.Count -gt 0) {
    $groupMembers | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
    Write-Host "  GroupMembers.csv              ($($groupMembers.Count) rows)" -ForegroundColor White
} else {
    [PSCustomObject][ordered]@{
        GroupId=''; GroupDisplayName=''; MemberId=''; MemberDisplayName=''
        MemberUPN='(empty - run with -ExpandGroups to populate)'
    } | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
    Write-Host '  GroupMembers.csv              (empty; re-run with -ExpandGroups)' -ForegroundColor DarkGray
}

# UnresolvedPrincipals.csv
$p = Join-Path $OutputFolder 'UnresolvedPrincipals.csv'
if ($unresolvedList.Count -gt 0) {
    $unresolvedList | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
} else {
    [PSCustomObject][ordered]@{
        PrincipalId=''; PrincipalType=''; ArtifactName=''; Source=''
    } | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
}
Write-Host "  UnresolvedPrincipals.csv      ($($unresolvedList.Count) rows)" -ForegroundColor White

# SecurityGroupSeed_<env>.csv  - Entra bulk-import format
#   Line 1 : version:v1.0  (literal - not a CSV header)
#   Line 2 : column header expected by the portal
#   Lines 3+: one UPN per line
$seedFileName = "SecurityGroupSeed_$envSafeName.csv"
$seedPath     = Join-Path $OutputFolder $seedFileName
$seedLines    = [System.Collections.Generic.List[string]]::new()
$seedLines.Add('version:v1.0')
$seedLines.Add('Member object ID or user principal name [memberObjectIdOrUpn] Required')
foreach ($upn in $sortedSeedUpns) { $seedLines.Add($upn) }
[System.IO.File]::WriteAllLines(
    $seedPath, $seedLines, [System.Text.Encoding]::UTF8)
Write-Host "  $seedFileName ($($sortedSeedUpns.Count) UPN(s))" -ForegroundColor White
Write-Host ''

###############################################################################
# 15.  COMPLETION BANNER
###############################################################################
Write-Host '  +======================================================================+' -ForegroundColor Cyan
Write-Host '  |  Audit complete.                                                      |' -ForegroundColor Cyan
Write-Host '  +======================================================================+' -ForegroundColor Cyan
Write-Host "  Environment : $envDisplayName" -ForegroundColor White
Write-Host "  Output      : $OutputFolder"   -ForegroundColor White
Write-Host ''
Write-Host '  IMPORTANT: Owners / share targets are *intended* users, not *proven* usage.' -ForegroundColor Yellow
Write-Host '  Cross-check with:' -ForegroundColor Yellow
Write-Host '    * Power Platform admin centre usage analytics' -ForegroundColor White
Write-Host '    * Dataverse audit logs' -ForegroundColor White
Write-Host "    * Microsoft 365 unified audit log ('Launched app' events)" -ForegroundColor White
Write-Host ''
Write-Host '  NEXT STEPS:' -ForegroundColor Cyan
Write-Host "  1. Review all CSVs with the business team." -ForegroundColor White
Write-Host "  2. Create an Entra ID security group for '$envDisplayName'." -ForegroundColor White
Write-Host "  3. Entra portal  >  Bulk import members  >  upload $seedFileName" -ForegroundColor White
Write-Host '  4. Power Platform admin centre  >  associate the group with this environment.' -ForegroundColor White
Write-Host '  5. Re-run this script with the next environment ID.' -ForegroundColor White
Write-Host ''
