<#
.SYNOPSIS
    Migrates API Manager contracts from a source API (in one Anypoint Business Group)
    to a destination API (in another Business Group / environment).

.DESCRIPTION
    - Authenticates to Anypoint Platform using a Connected App (Client ID / Secret) -> OAuth2 token.
    - Prompts the user to interactively select:
        * Environment
        * Source Business Group + Source API
        * Destination Business Group + Destination API
    - Reads all contracts (application <-> SLA tier bindings) on the source API.
    - For each contract:
        * If an equivalent contract already exists on the destination API (same application/client id),
          it is skipped and logged as "Contract already exists".
        * Otherwise, a new contract is created against the destination API and then approved,
          logged as "Contract created and approved".

.NOTES
    Requires PowerShell 5.1+ or PowerShell 7+.
    Anypoint APIs used:
        POST   /accounts/api/v2/oauth2/token
        GET    /accounts/api/organizations
        GET    /accounts/api/organizations/{orgId}/environments
        GET    /apimanager/api/v1/organizations/{orgId}/environments/{envId}/apis
        GET    /apimanager/api/v1/organizations/{orgId}/environments/{envId}/apis/{apiId}/policies... (n/a)
        GET    /apimanager/api/v1/organizations/{orgId}/environments/{envId}/apis/{apiId}/tiers
        GET    /apimanager/api/v1/organizations/{orgId}/environments/{envId}/apis/{apiId}/contracts
        POST   /apimanager/api/v1/organizations/{orgId}/environments/{envId}/apis/{apiId}/contracts
        POST   /apimanager/api/v1/organizations/{orgId}/environments/{envId}/apis/{apiId}/contracts/{contractId}/status
#>

# ============================================================
# CONFIGURATION
# ============================================================

$AnypointBaseUrl = "https://anypoint.mulesoft.com"

# ============================================================
# AUTH - Connected App (Client Credentials)
# ============================================================

function Get-AnypointToken {
    param(
        [Parameter(Mandatory)] [string]$ClientId,
        [Parameter(Mandatory)] [string]$ClientSecret
    )

    $uri = "$AnypointBaseUrl/accounts/api/v2/oauth2/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Error "Failed to obtain Anypoint token: $($_.Exception.Message)"
        throw
    }
}

# ============================================================
# GENERIC API WRAPPER
# ============================================================

function Invoke-AnypointApi {
    param(
        [Parameter(Mandatory)] [string]$Token,
        [Parameter(Mandatory)] [string]$Path,
        [string]$Method = "GET",
        $Body = $null
    )

    $uri = "$AnypointBaseUrl$Path"
    $headers = @{
        Authorization = "Bearer $Token"
        Accept        = "application/json"
    }

    $params = @{
        Uri     = $uri
        Method  = $Method
        Headers = $headers
    }

    if ($Body -ne $null) {
        $params["Body"] = ($Body | ConvertTo-Json -Depth 10)
        $params["ContentType"] = "application/json"
    }

    try {
        return Invoke-RestMethod @params
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $errBody = $null
        try { $errBody = ($_ | Get-Member -Name ErrorDetails -ErrorAction SilentlyContinue) ? $_.ErrorDetails.Message : $null } catch {}
        Write-Error "API call failed [$Method $Path] Status: $statusCode $errBody"
        throw
    }
}

# ============================================================
# GENERIC MENU SELECTOR
# ============================================================

function Select-FromList {
    param(
        [Parameter(Mandatory)] [array]$Items,
        [Parameter(Mandatory)] [string]$DisplayProperty,
        [Parameter(Mandatory)] [string]$Prompt
    )

    if ($Items.Count -eq 0) {
        Write-Error "No items available to select for: $Prompt"
        throw "Empty list"
    }

    Write-Host ""
    Write-Host "== $Prompt ==" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) {
        Write-Host "[$i] $($Items[$i].$DisplayProperty)"
    }

    do {
        $selection = Read-Host "Enter the number of your choice"
        $valid = ($selection -match '^\d+$') -and ([int]$selection -ge 0) -and ([int]$selection -lt $Items.Count)
        if (-not $valid) { Write-Host "Invalid selection, try again." -ForegroundColor Yellow }
    } until ($valid)

    return $Items[[int]$selection]
}

# ============================================================
# ANYPOINT DATA RETRIEVAL
# ============================================================

function Get-BusinessGroups {
    param([string]$Token)
    $result = Invoke-AnypointApi -Token $Token -Path "/accounts/api/organizations"
    return $result.data
}

function Get-Environments {
    param([string]$Token, [string]$OrgId)
    $result = Invoke-AnypointApi -Token $Token -Path "/accounts/api/organizations/$OrgId/environments"
    return $result.data
}

function Get-Apis {
    param([string]$Token, [string]$OrgId, [string]$EnvId)
    $result = Invoke-AnypointApi -Token $Token -Path "/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis"
    return $result.assets | ForEach-Object {
        foreach ($api in $_.apis) {
            [PSCustomObject]@{
                AssetId    = $_.assetId
                ApiId      = $api.id
                AssetName  = $_.assetId
                Version    = $api.assetVersion
                DisplayName = "$($_.assetId) (v$($api.assetVersion)) - apiId:$($api.id)"
            }
        }
    }
}

function Get-Contracts {
    param([string]$Token, [string]$OrgId, [string]$EnvId, [string]$ApiId)
    $result = Invoke-AnypointApi -Token $Token -Path "/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/contracts"
    return $result.contracts
}

function Get-SlaTiers {
    param([string]$Token, [string]$OrgId, [string]$EnvId, [string]$ApiId)
    $result = Invoke-AnypointApi -Token $Token -Path "/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/tiers"
    return $result.tiers
}

# ============================================================
# CONTRACT CREATION / APPROVAL
# ============================================================

function New-Contract {
    param(
        [string]$Token, [string]$OrgId, [string]$EnvId, [string]$ApiId,
        [string]$ApplicationId, [string]$TierId, [switch]$AcceptedTerms
    )

    $body = @{
        applicationId = $ApplicationId
        tierId        = $TierId
        acceptedTerms = $true
    }

    return Invoke-AnypointApi -Token $Token -Method "POST" `
        -Path "/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/contracts" `
        -Body $body
}

function Approve-Contract {
    param(
        [string]$Token, [string]$OrgId, [string]$EnvId, [string]$ApiId, [string]$ContractId
    )

    $body = @{ status = "APPROVED" }

    return Invoke-AnypointApi -Token $Token -Method "POST" `
        -Path "/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/contracts/$ContractId/status" `
        -Body $body
}

# ============================================================
# MAIN
# ============================================================

Write-Host "=== Anypoint Contract Migration Tool ===" -ForegroundColor Green

# --- Authenticate ---
$ClientId     = Read-Host "Enter Anypoint Connected App Client ID"
$ClientSecret = Read-Host "Enter Anypoint Connected App Client Secret" -AsSecureString
$ClientSecretPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
)

$Token = Get-AnypointToken -ClientId $ClientId -ClientSecret $ClientSecretPlain
Write-Host "Authenticated successfully." -ForegroundColor Green

# --- Business Groups ---
$AllBGs = Get-BusinessGroups -Token $Token

$SourceBG = Select-FromList -Items $AllBGs -DisplayProperty "name" -Prompt "Select SOURCE Business Group"
$DestBG   = Select-FromList -Items $AllBGs -DisplayProperty "name" -Prompt "Select DESTINATION Business Group"

# --- Environment (source) ---
$SourceEnvs = Get-Environments -Token $Token -OrgId $SourceBG.id
$SourceEnv  = Select-FromList -Items $SourceEnvs -DisplayProperty "name" -Prompt "Select Environment for SOURCE Business Group"

# --- Environment (destination) ---
$DestEnvs = Get-Environments -Token $Token -OrgId $DestBG.id
$DestEnv  = Select-FromList -Items $DestEnvs -DisplayProperty "name" -Prompt "Select Environment for DESTINATION Business Group"

# --- Source API ---
$SourceApis = Get-Apis -Token $Token -OrgId $SourceBG.id -EnvId $SourceEnv.id
$SourceApi  = Select-FromList -Items $SourceApis -DisplayProperty "DisplayName" -Prompt "Select SOURCE API"

# --- Destination API ---
$DestApis = Get-Apis -Token $Token -OrgId $DestBG.id -EnvId $DestEnv.id
$DestApi  = Select-FromList -Items $DestApis -DisplayProperty "DisplayName" -Prompt "Select DESTINATION API"

# --- Fetch contracts & tiers ---
Write-Host ""
Write-Host "Fetching contracts on source API..." -ForegroundColor Cyan
$SourceContracts = Get-Contracts -Token $Token -OrgId $SourceBG.id -EnvId $SourceEnv.id -ApiId $SourceApi.ApiId

Write-Host "Fetching contracts on destination API..." -ForegroundColor Cyan
$DestContracts = Get-Contracts -Token $Token -OrgId $DestBG.id -EnvId $DestEnv.id -ApiId $DestApi.ApiId

Write-Host "Fetching SLA tiers on destination API..." -ForegroundColor Cyan
$DestTiers = Get-SlaTiers -Token $Token -OrgId $DestBG.id -EnvId $DestEnv.id -ApiId $DestApi.ApiId

if ($SourceContracts.Count -eq 0) {
    Write-Host "No contracts found on source API. Nothing to migrate." -ForegroundColor Yellow
    return
}

# Build a lookup of destination contracts by applicationId for quick existence check
$DestContractsByApp = @{}
foreach ($c in $DestContracts) {
    $DestContractsByApp[$c.application.id] = $c
}

# Build a lookup of destination tiers by name (assumes same tier names should exist on destination)
$DestTiersByName = @{}
foreach ($t in $DestTiers) {
    $DestTiersByName[$t.name] = $t
}

# --- Migration Log ---
$Log = @()

foreach ($contract in $SourceContracts) {

    $appId       = $contract.application.id
    $appName     = $contract.application.name
    $sourceTier  = $contract.tier.name

    Write-Host ""
    Write-Host "Processing application: $appName (id: $appId), source tier: $sourceTier" -ForegroundColor White

    if ($DestContractsByApp.ContainsKey($appId)) {
        Write-Host "  -> Contract already exists on destination API." -ForegroundColor Yellow
        $Log += [PSCustomObject]@{
            Application = $appName
            ApplicationId = $appId
            SourceTier  = $sourceTier
            Status      = "Contract already exists"
        }
        continue
    }

    if (-not $DestTiersByName.ContainsKey($sourceTier)) {
        Write-Host "  -> No matching SLA tier '$sourceTier' found on destination API. Skipping." -ForegroundColor Red
        $Log += [PSCustomObject]@{
            Application = $appName
            ApplicationId = $appId
            SourceTier  = $sourceTier
            Status      = "Skipped - matching tier not found on destination"
        }
        continue
    }

    $destTierId = $DestTiersByName[$sourceTier].id

    try {
        $newContract = New-Contract -Token $Token -OrgId $DestBG.id -EnvId $DestEnv.id -ApiId $DestApi.ApiId `
            -ApplicationId $appId -TierId $destTierId

        Approve-Contract -Token $Token -OrgId $DestBG.id -EnvId $DestEnv.id -ApiId $DestApi.ApiId -ContractId $newContract.id

        Write-Host "  -> Contract created and approved." -ForegroundColor Green
        $Log += [PSCustomObject]@{
            Application = $appName
            ApplicationId = $appId
            SourceTier  = $sourceTier
            Status      = "Contract created and approved"
        }
    }
    catch {
        Write-Host "  -> Failed to create/approve contract: $($_.Exception.Message)" -ForegroundColor Red
        $Log += [PSCustomObject]@{
            Application = $appName
            ApplicationId = $appId
            SourceTier  = $sourceTier
            Status      = "Failed: $($_.Exception.Message)"
        }
    }
}

# --- Summary ---
Write-Host ""
Write-Host "=== Migration Summary ===" -ForegroundColor Green
$Log | Format-Table -AutoSize

$outFile = "AnypointContractMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$Log | Export-Csv -Path $outFile -NoTypeInformation
Write-Host "Summary exported to $outFile" -ForegroundColor Cyan
