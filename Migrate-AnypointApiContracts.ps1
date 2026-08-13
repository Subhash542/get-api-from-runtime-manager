#requires -Version 5.1
<#
    .SYNOPSIS
    Copies API contracts (client application access) from one API instance
    in a source Business Group to another API instance in a destination
    Business Group on Anypoint Platform.

    .DESCRIPTION
    Authenticates using a Connected App's Client ID / Client Secret (the
    supported way to call Anypoint APIs from a script) exchanged for a bearer
    access token, then prompts for:
      1. Source Business Group, source Environment, source API instance
      2. Destination Business Group, destination Environment, destination API instance

    A token copied out of a browser's DevTools Network tab is a short-lived
    UI session token and will not work reliably here - use a Connected App
    instead. See README.md for how to create one. If you already have a
    valid bearer token in hand (e.g. from your own token-issuing pipeline),
    you can skip the Client ID/Secret exchange with -Token.

    SLA tiers are migrated first: every SLA tier defined on the source API
    that doesn't already exist (matched by name) on the destination API is
    created there, copying its request limits. This mapping of source tier
    name -> destination tier id is then used when contracts are recreated,
    so a contract that used a given tier on the source API is attached to
    the equivalent tier on the destination API.

    For every contract found on the source API:
      - If a contract for the same consumer application already exists on the
        destination API -> logs "Contract already exists" and takes no action.
      - Otherwise, if a matching application (matched by Client ID, falling back
        to name) is found in the destination Business Group -> a new contract is
        created and immediately approved, logged as "Contract created and approved".
      - If no matching application exists in the destination Business Group, the
        contract is skipped and logged as "Application not found in destination -
        contract not created", since a contract cannot be created for an
        application that has not been registered in that Business Group.

    A summary is printed to the console and exported to a timestamped CSV file.

    .PARAMETER ClientId
    Client ID of an Anypoint Connected App (client_credentials grant). If
    omitted, and -Token is not supplied either, you'll be prompted for it.

    .PARAMETER ClientSecret
    Client Secret of the Connected App, as a SecureString. If omitted, and
    -Token is not supplied either, you'll be prompted for it (masked input).

    .PARAMETER Token
    Use this to supply an already-obtained bearer access token directly,
    bypassing the Client ID/Secret exchange. Must be a genuine API access
    token (e.g. from a Connected App), not a browser session token.

    .EXAMPLE
    .\Migrate-AnypointApiContracts.ps1
    Prompts for the Connected App Client ID and Client Secret interactively.

    .EXAMPLE
    .\Migrate-AnypointApiContracts.ps1 -ClientId 'abc123' -ClientSecret (Read-Host -AsSecureString)
    Supplies the Client ID up front and still masks the secret prompt.

    .NOTES
    Requires PowerShell 5.1+ (Windows PowerShell) or PowerShell 7+.
    Only reads from the source API; does not modify or revoke anything there.
#>

[CmdletBinding(DefaultParameterSetName = 'ConnectedApp')]
param(
    [Parameter(ParameterSetName = 'ConnectedApp')]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'ConnectedApp')]
    [SecureString]$ClientSecret,

    [Parameter(ParameterSetName = 'DirectToken')]
    [string]$Token,

    [string]$OutputFolder = (Join-Path -Path (Get-Location) -ChildPath 'AnypointContractMigrationResults')
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Modules\AnypointApi.psm1') -Force

function ConvertFrom-SecureStringPlain {
    param([Parameter(Mandatory)][SecureString]$Secure)
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Get-AnypointBearerToken {
    <#
        .SYNOPSIS
        Resolves the bearer token to use for this run: either the -Token
        override, or a Client ID/Secret exchange (prompting for whichever
        of ClientId/ClientSecret wasn't passed as a parameter).
    #>
    param(
        [string]$ClientId,
        [SecureString]$ClientSecret,
        [string]$Token
    )

    if ($Token) {
        return $Token.Trim()
    }

    if (-not $ClientId) {
        $ClientId = Read-Host -Prompt 'Enter the Connected App Client ID'
    }
    if (-not $ClientSecret) {
        $ClientSecret = Read-Host -Prompt 'Enter the Connected App Client Secret' -AsSecureString
    }
    $plainSecret = ConvertFrom-SecureStringPlain -Secure $ClientSecret

    Write-Host "Requesting an access token from the Connected App ..." -ForegroundColor DarkGray
    $tokenInfo = Get-AnypointAccessToken -ClientId $ClientId -ClientSecret $plainSecret
    Write-Host "Access token obtained (expires in $($tokenInfo.ExpiresIn)s)." -ForegroundColor DarkGray
    return $tokenInfo.AccessToken
}

function Select-BusinessGroupEnvironmentApi {
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][object[]]$Organizations,
        [Parameter(Mandatory)][string]$Label
    )

    $org = Select-FromList -Items $Organizations -DisplayProperty 'Path' -Prompt "Select the $Label Business Group"
    $environments = Get-AnypointEnvironments -Headers $Headers -OrgId $org.Id
    $env = Select-FromList -Items $environments -DisplayProperty 'Name' -Prompt "Select the $Label environment"

    Write-Host "Loading APIs for '$($org.Path)' / '$($env.Name)' ..." -ForegroundColor DarkGray
    $apis = Get-AnypointApiInstances -Headers $Headers -OrgId $org.Id -EnvId $env.Id
    if (-not $apis -or $apis.Count -eq 0) {
        throw "No API instances found in $Label Business Group / environment."
    }
    $apis = $apis | ForEach-Object {
        $_ | Add-Member -NotePropertyName Display -NotePropertyValue "$($_.Name) v$($_.AssetVersion) (instance $($_.InstanceLabel), id=$($_.ApiId))" -PassThru
    }
    $api = Select-FromList -Items $apis -DisplayProperty 'Display' -Prompt "Select the $Label API"

    return [pscustomobject]@{
        Org = $org
        Env = $env
        Api = $api
    }
}

# ---------------------------------------------------------------------------
# 1. Authenticate
# ---------------------------------------------------------------------------
Write-Host "=== Anypoint Platform Contract Migration ===" -ForegroundColor Green
$resolvedToken = Get-AnypointBearerToken -ClientId $ClientId -ClientSecret $ClientSecret -Token $Token
$headers = Get-AnypointAuthHeader -Token $resolvedToken

Write-Host "Validating token ..." -ForegroundColor DarkGray
Test-AnypointToken -Headers $headers | Out-Null

Write-Host "Loading Business Groups you have access to ..." -ForegroundColor DarkGray
$organizations = Get-AnypointOrganizations -Headers $headers
if (-not $organizations -or $organizations.Count -eq 0) {
    throw "No Business Groups were returned for this token. Verify the token is valid and not expired."
}

# ---------------------------------------------------------------------------
# 2. Source selection
# ---------------------------------------------------------------------------
$source = Select-BusinessGroupEnvironmentApi -Headers $headers -Organizations $organizations -Label 'SOURCE'

# ---------------------------------------------------------------------------
# 3. Destination selection
# ---------------------------------------------------------------------------
$destination = Select-BusinessGroupEnvironmentApi -Headers $headers -Organizations $organizations -Label 'DESTINATION'

if ($source.Api.ApiId -eq $destination.Api.ApiId -and $source.Env.Id -eq $destination.Env.Id) {
    throw "Source and destination API instance are the same. Nothing to migrate."
}

# ---------------------------------------------------------------------------
# 4. Load contracts / applications
# ---------------------------------------------------------------------------
Write-Host "`nLoading contracts on source API '$($source.Api.Name)' ..." -ForegroundColor DarkGray
$sourceContracts = Get-AnypointContracts -Headers $headers -OrgId $source.Org.Id -EnvId $source.Env.Id -ApiId $source.Api.ApiId
Write-Host "Found $($sourceContracts.Count) contract(s) on the source API."

Write-Host "Loading existing contracts on destination API '$($destination.Api.Name)' ..." -ForegroundColor DarkGray
$destinationContracts = Get-AnypointContracts -Headers $headers -OrgId $destination.Org.Id -EnvId $destination.Env.Id -ApiId $destination.Api.ApiId

Write-Host "Loading applications registered in destination Business Group ..." -ForegroundColor DarkGray
$destinationApps = Get-AnypointApplications -Headers $headers -OrgId $destination.Org.Id -EnvId $destination.Env.Id

# ---------------------------------------------------------------------------
# 5. Migrate SLA tiers (by name) before touching contracts, so contract
#    creation below can map a source tier id to the right destination tier id.
# ---------------------------------------------------------------------------
Write-Host "`nLoading SLA tiers on source API ..." -ForegroundColor DarkGray
$sourceTiers = Get-AnypointTiers -Headers $headers -OrgId $source.Org.Id -EnvId $source.Env.Id -ApiId $source.Api.ApiId
Write-Host "Loading SLA tiers on destination API ..." -ForegroundColor DarkGray
$destinationTiers = Get-AnypointTiers -Headers $headers -OrgId $destination.Org.Id -EnvId $destination.Env.Id -ApiId $destination.Api.ApiId

$tierResults = New-Object System.Collections.Generic.List[object]
# Maps a source tier id -> the tier id to use on the destination API.
$tierIdMap = @{}

foreach ($tier in $sourceTiers) {
    $existingTier = $destinationTiers | Where-Object { $_.Name -eq $tier.Name } | Select-Object -First 1

    if ($existingTier) {
        $tierIdMap[$tier.TierId] = $existingTier.TierId
        $tierResults.Add([pscustomobject]@{
            Tier    = $tier.Name
            Action  = 'None'
            Comment = 'Tier already exists on destination API'
        })
        continue
    }

    try {
        $newTier = New-AnypointTier -Headers $headers -OrgId $destination.Org.Id -EnvId $destination.Env.Id `
            -ApiId $destination.Api.ApiId -SourceTier $tier
        $tierIdMap[$tier.TierId] = $newTier.id
        $tierResults.Add([pscustomobject]@{
            Tier    = $tier.Name
            Action  = 'Created'
            Comment = 'Tier created on destination API'
        })
    }
    catch {
        $tierResults.Add([pscustomobject]@{
            Tier    = $tier.Name
            Action  = 'Error'
            Comment = "Failed to create tier: $_"
        })
    }
}

if ($sourceTiers.Count -gt 0) {
    Write-Host "`n=== SLA Tier Migration Summary ===" -ForegroundColor Green
    $tierResults | Format-Table -AutoSize
}

# ---------------------------------------------------------------------------
# 6. Process each source contract
# ---------------------------------------------------------------------------
$results = New-Object System.Collections.Generic.List[object]

foreach ($contract in $sourceContracts) {

    if ($contract.Status -eq 'REVOKED') {
        $results.Add([pscustomobject]@{
            Application = $contract.AppName
            ClientId    = $contract.ClientId
            Action      = 'Skipped'
            Comment     = 'Source contract is revoked - not migrated'
        })
        continue
    }

    # a) Does a contract for this same consumer already exist on the destination API?
    $existing = $destinationContracts | Where-Object {
        ($_.ClientId -and $contract.ClientId -and $_.ClientId -eq $contract.ClientId) -or
        ($_.AppName -eq $contract.AppName)
    } | Select-Object -First 1

    if ($existing) {
        $results.Add([pscustomobject]@{
            Application = $contract.AppName
            ClientId    = $contract.ClientId
            Action      = 'None'
            Comment     = 'Contract already exists'
        })
        continue
    }

    # b) Find the matching application registered in the destination Business Group
    $matchedApp = $destinationApps | Where-Object {
        ($_.ClientId -and $contract.ClientId -and $_.ClientId -eq $contract.ClientId) -or
        ($_.Name -eq $contract.AppName)
    } | Select-Object -First 1

    if (-not $matchedApp) {
        $results.Add([pscustomobject]@{
            Application = $contract.AppName
            ClientId    = $contract.ClientId
            Action      = 'Skipped'
            Comment     = 'Application not found in destination Business Group - register it there first'
        })
        continue
    }

    # c) Create + approve the contract, mapping the source tier id to its
    #    equivalent (existing or newly-created) tier id on the destination API.
    $destinationTierId = $null
    if ($contract.TierId) {
        $destinationTierId = $tierIdMap[$contract.TierId]
        if (-not $destinationTierId) {
            Write-Warning "No destination tier mapping found for source tier id $($contract.TierId) (application $($contract.AppName)) - creating contract without a tier."
        }
    }

    try {
        $created = New-AnypointContract -Headers $headers -OrgId $destination.Org.Id -EnvId $destination.Env.Id `
            -ApiId $destination.Api.ApiId -ApplicationId $matchedApp.ApplicationId -TierId $destinationTierId

        Approve-AnypointContract -Headers $headers -OrgId $destination.Org.Id -EnvId $destination.Env.Id `
            -ApiId $destination.Api.ApiId -ContractId $created.id

        $results.Add([pscustomobject]@{
            Application = $contract.AppName
            ClientId    = $contract.ClientId
            Action      = 'Created'
            Comment     = 'Contract created and approved'
        })
    }
    catch {
        $results.Add([pscustomobject]@{
            Application = $contract.AppName
            ClientId    = $contract.ClientId
            Action      = 'Error'
            Comment     = "Failed to create/approve contract: $_"
        })
    }
}

# ---------------------------------------------------------------------------
# 7. Report
# ---------------------------------------------------------------------------
Write-Host "`n=== Contract Migration Summary ===" -ForegroundColor Green
$results | Format-Table -AutoSize

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

$csvPath = Join-Path $OutputFolder ("ContractMigration_{0}.csv" -f $timestamp)
$results | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "`nContract results exported to: $csvPath" -ForegroundColor Cyan

if ($tierResults.Count -gt 0) {
    $tierCsvPath = Join-Path $OutputFolder ("TierMigration_{0}.csv" -f $timestamp)
    $tierResults | Export-Csv -Path $tierCsvPath -NoTypeInformation
    Write-Host "Tier results exported to: $tierCsvPath" -ForegroundColor Cyan
}
