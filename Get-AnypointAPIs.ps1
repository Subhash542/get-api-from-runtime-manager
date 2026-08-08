<#
.SYNOPSIS
    Lists all APIs/applications (with status and deployment details) from BOTH
    API Manager and Runtime Manager in Anypoint Platform, across ALL business
    groups, for a user-selected environment.

.DESCRIPTION
    - Authenticates using a Bearer token (Connected App token or personal access token)
      that you already have — the script does NOT perform the OAuth login itself.
    - Enumerates every business group (org + nested sub-orgs) the token has access to.
    - Collects the distinct set of environments across those business groups and
      prompts the user to pick one.
    - Queries API Manager for every business group that has that environment and
      returns Asset Name/ID, Version, Status, Instance Type, Deployment Type, etc.
    - Queries Runtime Manager for every deployed application in that environment
      via TWO APIs, since your org uses both CloudHub 2.0 and on-prem Hybrid:
        (a) Application Manager API (/amc/application-manager/...) — covers
            CloudHub 2.0 and Runtime Fabric deployments.
        (b) ARM / Hybrid API (/hybrid/api/v1/...) — covers on-prem/Hybrid
            standalone and clustered Mule servers registered with the
            Runtime Manager agent.
      Results from both are merged into one list, tagged with a Source column.
    - Optionally exports both result sets to CSV.

.PARAMETER Token
    Anypoint Platform bearer access token. If omitted, you'll be prompted (masked input).

.PARAMETER BaseUri
    Anypoint control-plane base URL. Defaults to https://anypoint.mulesoft.com
    (use https://eu1.anypoint.mulesoft.com for EU control plane).

.PARAMETER ExportCsv
    Optional path to export API Manager results as CSV, e.g. -ExportCsv .\apis.csv

.PARAMETER ExportRuntimeCsv
    Optional path to export Runtime Manager application results as CSV,
    e.g. -ExportRuntimeCsv .\runtime-apps.csv

.EXAMPLE
    .\Get-AnypointAPIs.ps1

.EXAMPLE
    .\Get-AnypointAPIs.ps1 -Token "eyJhbGciOi..." -ExportCsv .\apis.csv -ExportRuntimeCsv .\runtime-apps.csv

.NOTES
    Required token scopes: View Environment / API Manager / Runtime Manager read
    access on the business groups you want to query.

    Anypoint exposes Runtime Manager data through more than one API depending on
    deployment target, and exact field names can shift slightly between API
    versions/org configurations:
      - CloudHub 2.0 / RTF -> /amc/application-manager/api/v2/.../deployments
      - Hybrid / on-prem   -> /hybrid/api/v1/applications (org+env via headers)
    This script queries both and merges the results. If a column comes back
    blank for your tenant, uncomment the debug dump line noted in each section
    to inspect the raw JSON and adjust the property path.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Token,

    [Parameter(Mandatory = $false)]
    [string]$BaseUri = "https://anypoint.mulesoft.com",

    [Parameter(Mandatory = $false)]
    [string]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [string]$ExportRuntimeCsv
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# 0. Get token if not supplied
# ---------------------------------------------------------------------------
if (-not $Token) {
    $secureToken = Read-Host -Prompt "Enter Anypoint Platform Bearer token" -AsSecureString
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken)
    $Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}

$Headers = @{
    "Authorization" = "Bearer $Token"
    "Content-Type"  = "application/json"
}

function Invoke-AnypointGet {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [hashtable]$QueryParams,
        [hashtable]$ExtraHeaders
    )
    $uri = "$BaseUri$Path"
    if ($QueryParams -and $QueryParams.Count -gt 0) {
        $qs = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
        $uri = "$uri`?$qs"
    }
    $reqHeaders = $Headers.Clone()
    if ($ExtraHeaders) {
        foreach ($k in $ExtraHeaders.Keys) { $reqHeaders[$k] = $ExtraHeaders[$k] }
    }
    try {
        return Invoke-RestMethod -Uri $uri -Headers $reqHeaders -Method Get
    }
    catch {
        $status = $_.Exception.Response.StatusCode.value__
        Write-Warning "GET $uri failed (HTTP $status): $($_.Exception.Message)"
        return $null
    }
}

# Defensively pulls the first present property from a list of dotted paths,
# since Anypoint's Runtime Manager response shape varies by API/target type.
function Resolve-Prop {
    param(
        [Parameter(Mandatory)] $Object,
        [Parameter(Mandatory)] [string[]]$Paths
    )
    foreach ($path in $Paths) {
        $current = $Object
        $found = $true
        foreach ($segment in $path.Split('.')) {
            if ($null -ne $current -and $current.PSObject.Properties.Name -contains $segment) {
                $current = $current.$segment
            }
            else {
                $found = $false
                break
            }
        }
        if ($found -and $null -ne $current) { return $current }
    }
    return $null
}

# ---------------------------------------------------------------------------
# 1. Verify token / get identity
# ---------------------------------------------------------------------------
Write-Host "Validating token..." -ForegroundColor Cyan
$me = Invoke-AnypointGet -Path "/accounts/api/me"
if (-not $me) {
    Write-Error "Could not authenticate with the supplied token. Aborting."
    exit 1
}
Write-Host "Authenticated as: $($me.user.username)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 2. Recursively enumerate all business groups (org + sub-orgs)
# ---------------------------------------------------------------------------
function Get-AllBusinessGroups {
    param([string]$RootOrgId)

    $allOrgs = @()
    $toProcess = [System.Collections.Generic.Queue[string]]::new()
    $toProcess.Enqueue($RootOrgId)
    $seen = @{}

    while ($toProcess.Count -gt 0) {
        $orgId = $toProcess.Dequeue()
        if ($seen.ContainsKey($orgId)) { continue }
        $seen[$orgId] = $true

        $org = Invoke-AnypointGet -Path "/accounts/api/organizations/$orgId"
        if (-not $org) { continue }

        $allOrgs += [PSCustomObject]@{
            Id   = $org.id
            Name = $org.name
        }

        foreach ($subId in $org.subOrganizationIds) {
            if (-not $seen.ContainsKey($subId)) {
                $toProcess.Enqueue($subId)
            }
        }
    }
    return $allOrgs
}

Write-Host "Discovering business groups..." -ForegroundColor Cyan
$rootOrgId = $me.user.organization.id
$businessGroups = Get-AllBusinessGroups -RootOrgId $rootOrgId
Write-Host "Found $($businessGroups.Count) business group(s):" -ForegroundColor Green
$businessGroups | ForEach-Object { Write-Host "  - $($_.Name) [$($_.Id)]" }

if ($businessGroups.Count -eq 0) {
    Write-Error "No business groups found for this token."
    exit 1
}

# ---------------------------------------------------------------------------
# 3. Enumerate environments per business group, build distinct env-name list
# ---------------------------------------------------------------------------
Write-Host "`nDiscovering environments across business groups..." -ForegroundColor Cyan

# Map: envName -> list of @{ OrgId, OrgName, EnvId }
$envMap = @{}

foreach ($bg in $businessGroups) {
    $envResp = Invoke-AnypointGet -Path "/accounts/api/organizations/$($bg.Id)/environments"
    if (-not $envResp) { continue }

    foreach ($env in $envResp.data) {
        if (-not $envMap.ContainsKey($env.name)) {
            $envMap[$env.name] = @()
        }
        $envMap[$env.name] += [PSCustomObject]@{
            OrgId   = $bg.Id
            OrgName = $bg.Name
            EnvId   = $env.id
            EnvType = $env.type
        }
    }
}

if ($envMap.Count -eq 0) {
    Write-Error "No environments found across any business group."
    exit 1
}

# ---------------------------------------------------------------------------
# 4. Prompt user to select an environment (by name)
# ---------------------------------------------------------------------------
$envNames = $envMap.Keys | Sort-Object
Write-Host "`nAvailable environments:" -ForegroundColor Cyan
for ($i = 0; $i -lt $envNames.Count; $i++) {
    $matchCount = $envMap[$envNames[$i]].Count
    Write-Host "  [$($i+1)] $($envNames[$i])  (present in $matchCount business group(s))"
}

$selection = $null
while (-not $selection) {
    $inputVal = Read-Host "`nSelect an environment by number (1-$($envNames.Count))"
    if ($inputVal -match '^\d+$' -and [int]$inputVal -ge 1 -and [int]$inputVal -le $envNames.Count) {
        $selection = $envNames[[int]$inputVal - 1]
    }
    else {
        Write-Warning "Invalid selection. Try again."
    }
}

Write-Host "Selected environment: $selection" -ForegroundColor Green
$targets = $envMap[$selection]

# ---------------------------------------------------------------------------
# 5. Query API Manager for APIs in the selected environment, per business group
# ---------------------------------------------------------------------------
Write-Host "`nFetching APIs from API/Runtime Manager..." -ForegroundColor Cyan

$results = New-Object System.Collections.Generic.List[Object]

foreach ($target in $targets) {

    $offset = 0
    $limit  = 50
    $total  = $null

    do {
        $apiPath = "/apimanager/api/v1/organizations/$($target.OrgId)/environments/$($target.EnvId)/apis"
        $resp = Invoke-AnypointGet -Path $apiPath -QueryParams @{ limit = $limit; offset = $offset }
        if (-not $resp) { break }

        if ($null -eq $total) { $total = $resp.total }

        foreach ($api in $resp.assets) {
            # Each "asset" can have multiple API versions/instances
            foreach ($instance in $api.apis) {
                $results.Add([PSCustomObject]@{
                    BusinessGroup   = $target.OrgName
                    Environment     = $selection
                    AssetId         = $api.groupId + ":" + $api.assetId
                    AssetName       = $api.assetId
                    ApiId           = $instance.id
                    AssetVersion    = $instance.assetVersion
                    ProductVersion  = $instance.productVersion
                    InstanceLabel   = $instance.instanceLabel
                    Status          = $instance.status
                    IsDeployed      = $instance.isDeployed
                    DeploymentType  = $instance.deployment.type
                    RuntimeVersion  = $instance.deployment.gatewayVersion
                    Technology      = $instance.technology
                    LastActiveDate  = $instance.lastActiveDate
                })
            }
        }

        $offset += $limit
    } while ($total -and $offset -lt $total)
}

# ---------------------------------------------------------------------------
# 6a. Runtime Manager — CloudHub 2.0 / RTF via Application Manager API
# ---------------------------------------------------------------------------
Write-Host "`nFetching CloudHub 2.0 / RTF applications from Runtime Manager..." -ForegroundColor Cyan

$runtimeResults = New-Object System.Collections.Generic.List[Object]

foreach ($target in $targets) {

    $offset = 0
    $limit  = 50
    $total  = $null

    do {
        $rtPath = "/amc/application-manager/api/v2/organizations/$($target.OrgId)/environments/$($target.EnvId)/deployments"
        $resp = Invoke-AnypointGet -Path $rtPath -QueryParams @{ limit = $limit; offset = $offset }
        if (-not $resp) { break }

        # Uncomment to inspect the raw shape if fields below come back empty:
        # $resp | ConvertTo-Json -Depth 10 | Out-File .\_debug_amc_response.json -Append

        if ($null -eq $total) { $total = $resp.total }

        foreach ($dep in $resp.data) {
            $runtimeResults.Add([PSCustomObject]@{
                Source           = "CloudHub 2.0 / RTF"
                BusinessGroup    = $target.OrgName
                Environment      = $selection
                ApplicationName  = $dep.name
                Status           = Resolve-Prop $dep @("status", "application.status", "lastReportedStatus")
                TargetType       = Resolve-Prop $dep @("target.type")
                TargetName       = Resolve-Prop $dep @("target.name")
                MuleVersion      = Resolve-Prop $dep @("application.runtimeVersion", "target.deploymentSettings.runtimeVersion")
                Replicas         = Resolve-Prop $dep @("target.deploymentSettings.instances", "target.deploymentSettings.replicas")
                VCores           = Resolve-Prop $dep @("application.vCores", "target.deploymentSettings.resources.cpu.reserved")
                LastUpdated      = Resolve-Prop $dep @("lastModifiedDate", "lastReportedSync")
                ApplicationId    = $dep.id
            })
        }

        $offset += $limit
    } while ($total -and $offset -lt $total)
}

# ---------------------------------------------------------------------------
# 6b. Runtime Manager — Hybrid / on-prem via ARM (Anypoint Runtime Manager) API
#     ARM is scoped via headers (X-ANYPNT-ORG-ID / X-ANYPNT-ENV-ID), not the URL path.
# ---------------------------------------------------------------------------
Write-Host "Fetching Hybrid / on-prem applications from Runtime Manager (ARM)..." -ForegroundColor Cyan

foreach ($target in $targets) {

    $armHeaders = @{
        "X-ANYPNT-ORG-ID" = $target.OrgId
        "X-ANYPNT-ENV-ID" = $target.EnvId
    }

    $resp = Invoke-AnypointGet -Path "/hybrid/api/v1/applications" -ExtraHeaders $armHeaders
    if (-not $resp) { continue }

    # Uncomment to inspect the raw shape if fields below come back empty:
    # $resp | ConvertTo-Json -Depth 10 | Out-File .\_debug_arm_response.json -Append

    $apps = if ($resp.data) { $resp.data } else { $resp }

    foreach ($app in $apps) {
        $runtimeResults.Add([PSCustomObject]@{
            Source           = "Hybrid / On-Prem"
            BusinessGroup    = $target.OrgName
            Environment      = $selection
            ApplicationName  = Resolve-Prop $app @("name", "artifact.name")
            Status           = Resolve-Prop $app @("lastReportedStatus", "artifact.status", "status", "desiredStatus")
            TargetType       = Resolve-Prop $app @("target.type")
            TargetName       = Resolve-Prop $app @("target.name", "server.name")
            MuleVersion      = Resolve-Prop $app @("muleVersion", "target.details.muleVersion", "artifact.muleVersion")
            Replicas         = $null
            VCores           = $null
            LastUpdated      = Resolve-Prop $app @("lastReportedStatusUpdateTime", "lastModifiedDate")
            ApplicationId    = $app.id
        })
    }
}

# ---------------------------------------------------------------------------
# 7. Output
# ---------------------------------------------------------------------------
if ($results.Count -eq 0) {
    Write-Host "`nNo APIs found in API Manager for environment '$selection' across the discovered business groups." -ForegroundColor Yellow
}
else {
    Write-Host "`n=== API Manager: $($results.Count) API instance(s) ===" -ForegroundColor Green
    $results | Sort-Object BusinessGroup, AssetName |
        Format-Table BusinessGroup, AssetName, AssetVersion, Status, DeploymentType, RuntimeVersion, IsDeployed -AutoSize

    if ($ExportCsv) {
        $results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
        Write-Host "Exported API Manager details to: $ExportCsv" -ForegroundColor Green
    }
}

if ($runtimeResults.Count -eq 0) {
    Write-Host "`nNo applications found in Runtime Manager for environment '$selection' across the discovered business groups." -ForegroundColor Yellow
}
else {
    Write-Host "`n=== Runtime Manager: $($runtimeResults.Count) application(s) ===" -ForegroundColor Green
    $runtimeResults | Sort-Object Source, BusinessGroup, ApplicationName |
        Format-Table Source, BusinessGroup, ApplicationName, Status, TargetType, TargetName, MuleVersion, LastUpdated -AutoSize

    if ($ExportRuntimeCsv) {
        $runtimeResults | Export-Csv -Path $ExportRuntimeCsv -NoTypeInformation -Encoding UTF8
        Write-Host "Exported Runtime Manager details to: $ExportRuntimeCsv" -ForegroundColor Green
    }
}
