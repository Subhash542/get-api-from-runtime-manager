#requires -Version 5.1
<#
    AnypointApi.psm1
    -----------------
    Helper functions that wrap the Anypoint Platform REST APIs used to:
      - list Business Groups (organizations / sub-organizations)
      - list Environments
      - list API Manager API instances
      - list / create / approve API contracts
      - list client applications (used to match consumers across Business Groups)

    All calls use Bearer token authentication. The supported way to get that
    token is via a Connected App's Client ID / Client Secret exchanged for an
    access token (see Get-AnypointAccessToken below) — not a token copied out
    of a browser's DevTools, which is a short-lived UI session token and will
    not reliably work for scripted API calls.

    NOTE ON API SURFACE:
    Anypoint's Platform APIs are versioned per-tenant/per-release. The endpoints
    below reflect the documented API Manager v1 / Access Management APIs at the
    time this script was written. If your organization is on a different
    API Manager version (e.g. still using the legacy /apiplatform/repository/v2
    APIs) you may need to adjust the URIs in this file. Everything that talks
    to Anypoint is isolated in this module so that adjustment only needs to
    happen in one place.
#>

Set-StrictMode -Version Latest
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$script:AnypointBaseUri = 'https://anypoint.mulesoft.com'

function Get-AnypointAccessToken {
    <#
        .SYNOPSIS
        Exchanges a Connected App's Client ID / Client Secret for a bearer
        access token via the client_credentials grant. This is the
        supported way to authenticate a script against Anypoint Platform
        APIs (a token copied out of the browser's DevTools is a UI session
        token and is not fit for this purpose - it expires in minutes and
        may be scoped only for the frontend app).
    #>
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $body = @{
        client_id     = $ClientId.Trim()
        client_secret = $ClientSecret.Trim()
        grant_type    = 'client_credentials'
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri "$script:AnypointBaseUri/accounts/api/v2/oauth2/token" `
            -Method POST -ContentType 'application/json' -Body $body -ErrorAction Stop
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message
        $msg = "Failed to obtain an access token from the Connected App credentials."
        if ($errorDetail) { $msg += " $errorDetail" } else { $msg += " $($_.Exception.Message)" }
        $msg += " Double-check the Client ID/Secret and that the Connected App is enabled with the 'client_credentials' grant."
        throw $msg
    }

    if (-not $response.access_token) {
        throw "Token endpoint returned no access_token. Response: $($response | ConvertTo-Json -Depth 5)"
    }

    return [pscustomobject]@{
        AccessToken = $response.access_token
        ExpiresIn   = $response.expires_in
        ObtainedAt  = Get-Date
    }
}

function Get-AnypointAuthHeader {
    <#
        .SYNOPSIS
        Builds the Authorization header dictionary from a bearer token.
    #>
    param(
        [Parameter(Mandatory)][string]$Token
    )
    return @{
        'Authorization' = "Bearer $($Token.Trim())"
        'Content-Type'  = 'application/json'
    }
}

function Test-AnypointToken {
    <#
        .SYNOPSIS
        Validates a token before doing anything else, so a bad/expired token
        fails fast with a clear message instead of surfacing as a generic
        401 deep into the script.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers
    )
    try {
        Invoke-AnypointApi -Headers $Headers -Uri "$script:AnypointBaseUri/accounts/api/profile" | Out-Null
        return $true
    }
    catch {
        if ($_ -match 'HTTP 401') {
            throw "Token was rejected (HTTP 401). This usually means: the token is expired, it's the wrong type (Personal Access Tokens are deprecated on most orgs - use a Connected App access token instead), or it was pasted with extra whitespace. Get a fresh token and try again."
        }
        throw
    }
}

function Invoke-AnypointApi {
    <#
        .SYNOPSIS
        Thin wrapper around Invoke-RestMethod with consistent error handling.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Headers,
        [ValidateSet('GET','POST','PATCH','PUT','DELETE')][string]$Method = 'GET',
        [object]$Body = $null
    )

    try {
        $params = @{
            Uri     = $Uri
            Headers = $Headers
            Method  = $Method
            ErrorAction = 'Stop'
        }
        if ($null -ne $Body) {
            $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
        }
        return Invoke-RestMethod @params
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message
        $status = $null
        if ($_.Exception.Response) {
            try { $status = [int]$_.Exception.Response.StatusCode } catch {}
        }
        $msg = "Anypoint API call failed [$Method $Uri]"
        if ($status) { $msg += " (HTTP $status)" }
        if ($errorDetail) { $msg += ": $errorDetail" } else { $msg += ": $($_.Exception.Message)" }
        throw $msg
    }
}

function Get-AnypointOrganizations {
    <#
        .SYNOPSIS
        Returns a flattened list of Business Groups (the root organization and
        every nested sub-organization) the authenticated user can see.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers
    )

    $result = Invoke-AnypointApi -Headers $Headers -Uri "$script:AnypointBaseUri/accounts/api/organizations"

    $flat = New-Object System.Collections.Generic.List[object]

    function Add-OrgRecursive($org, $parentPath) {
        $path = if ($parentPath) { "$parentPath > $($org.name)" } else { $org.name }
        $flat.Add([pscustomobject]@{
            Id   = $org.id
            Name = $org.name
            Path = $path
        })
        if ($org.subOrganizationIds -and $org.subOrganizationIds.Count -gt 0) {
            foreach ($subId in $org.subOrganizationIds) {
                try {
                    $sub = Invoke-AnypointApi -Headers $Headers -Uri "$script:AnypointBaseUri/accounts/api/organizations/$subId"
                    Add-OrgRecursive -org $sub -parentPath $path
                } catch {
                    Write-Warning "Could not load sub-organization $subId : $_"
                }
            }
        }
    }

    foreach ($org in $result.organizations) {
        Add-OrgRecursive -org $org -parentPath $null
    }

    return $flat | Sort-Object Path
}

function Get-AnypointEnvironments {
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId
    )
    $result = Invoke-AnypointApi -Headers $Headers -Uri "$script:AnypointBaseUri/accounts/api/organizations/$OrgId/environments"
    return $result.data | ForEach-Object {
        [pscustomobject]@{ Id = $_.id; Name = $_.name; Type = $_.type }
    } | Sort-Object Name
}

function Get-AnypointApiInstances {
    <#
        .SYNOPSIS
        Lists API Manager API instances (assets) for a given org/environment.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId,
        [Parameter(Mandatory)][string]$EnvId
    )

    $all = New-Object System.Collections.Generic.List[object]
    $offset = 0
    $limit  = 50

    do {
        $uri = "$script:AnypointBaseUri/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis?offset=$offset&limit=$limit"
        $result = Invoke-AnypointApi -Headers $Headers -Uri $uri

        foreach ($asset in $result.assets) {
            foreach ($api in $asset.apis) {
                $all.Add([pscustomobject]@{
                    ApiId         = $api.id
                    AssetId       = $asset.assetId
                    Name          = $asset.name
                    AssetVersion  = $api.assetVersion
                    InstanceLabel = $api.instanceLabel
                    Technology    = $api.technology
                    EnvironmentId = $EnvId
                    OrgId         = $OrgId
                })
            }
        }

        $offset += $limit
    } while ($all.Count -lt $result.total)

    return $all
}

function Get-AnypointContracts {
    <#
        .SYNOPSIS
        Lists all contracts for a given API instance.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId,
        [Parameter(Mandatory)][string]$EnvId,
        [Parameter(Mandatory)][string]$ApiId
    )

    $all = New-Object System.Collections.Generic.List[object]
    $offset = 0
    $limit  = 50

    do {
        $uri = "$script:AnypointBaseUri/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/contracts?offset=$offset&limit=$limit"
        $result = Invoke-AnypointApi -Headers $Headers -Uri $uri

        foreach ($c in $result.contracts) {
            $all.Add([pscustomobject]@{
                ContractId    = $c.id
                ApplicationId = $c.applicationId
                AppName       = $c.application.name
                ClientId      = $c.application.coreServicesId
                Status        = $c.status
                TierId        = $c.tierId
                AcceptedTerms = $c.acceptedTerms
            })
        }

        $offset += $limit
    } while ($all.Count -lt $result.total)

    return $all
}

function Get-AnypointApplications {
    <#
        .SYNOPSIS
        Lists client applications registered under an org/environment
        (used to match a source contract's consumer to a destination application).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId,
        [Parameter(Mandatory)][string]$EnvId
    )

    $all = New-Object System.Collections.Generic.List[object]
    $offset = 0
    $limit  = 50

    do {
        $uri = "$script:AnypointBaseUri/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/applications?offset=$offset&limit=$limit"
        $result = Invoke-AnypointApi -Headers $Headers -Uri $uri

        $items = if ($result.applications) { $result.applications } else { $result }
        foreach ($a in $items) {
            $all.Add([pscustomobject]@{
                ApplicationId = $a.id
                Name          = $a.name
                ClientId      = $a.coreServicesId
            })
        }

        if ($result.total) { $offset += $limit } else { break }
    } while ($result.total -and $all.Count -lt $result.total)

    return $all
}

function Get-AnypointTiers {
    <#
        .SYNOPSIS
        Lists SLA tiers defined on an API instance.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId,
        [Parameter(Mandatory)][string]$EnvId,
        [Parameter(Mandatory)][string]$ApiId
    )

    $all = New-Object System.Collections.Generic.List[object]
    $offset = 0
    $limit  = 50

    do {
        $uri = "$script:AnypointBaseUri/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/tiers?offset=$offset&limit=$limit"
        $result = Invoke-AnypointApi -Headers $Headers -Uri $uri

        $items = if ($null -ne $result.tiers) { $result.tiers } else { $result }
        foreach ($t in $items) {
            $all.Add([pscustomobject]@{
                TierId               = $t.id
                Name                 = $t.name
                Description          = $t.description
                Status               = $t.status
                ApplicationsCount    = $t.applicationsCount
                Limits               = $t.limits
            })
        }

        if ($result.total) { $offset += $limit } else { break }
    } while ($result.total -and $all.Count -lt $result.total)

    return $all
}

function New-AnypointTier {
    <#
        .SYNOPSIS
        Creates a new SLA tier on an API instance, copying name/description/limits
        from a source tier object (as returned by Get-AnypointTiers).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId,
        [Parameter(Mandatory)][string]$EnvId,
        [Parameter(Mandatory)][string]$ApiId,
        [Parameter(Mandatory)][object]$SourceTier
    )

    $limits = @()
    foreach ($l in $SourceTier.Limits) {
        $limits += @{
            maximumRequests          = $l.maximumRequests
            timePeriodInMilliseconds = $l.timePeriodInMilliseconds
            visible                  = if ($null -ne $l.visible) { $l.visible } else { $true }
        }
    }

    $body = @{
        name        = $SourceTier.Name
        description = $SourceTier.Description
        status      = if ($SourceTier.Status) { $SourceTier.Status } else { 'ACTIVE' }
        limits      = $limits
    }

    $uri = "$script:AnypointBaseUri/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/tiers"
    return Invoke-AnypointApi -Headers $Headers -Uri $uri -Method POST -Body $body
}

function New-AnypointContract {
    <#
        .SYNOPSIS
        Creates a new (pending) contract between an application and an API instance.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId,
        [Parameter(Mandatory)][string]$EnvId,
        [Parameter(Mandatory)][string]$ApiId,
        [Parameter(Mandatory)][string]$ApplicationId,
        [string]$TierId
    )

    $body = @{
        applicationId = $ApplicationId
        acceptedTerms = $true
    }
    if ($TierId) { $body['tierId'] = $TierId }

    $uri = "$script:AnypointBaseUri/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/contracts"
    return Invoke-AnypointApi -Headers $Headers -Uri $uri -Method POST -Body $body
}

function Approve-AnypointContract {
    <#
        .SYNOPSIS
        Approves a pending contract.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$OrgId,
        [Parameter(Mandatory)][string]$EnvId,
        [Parameter(Mandatory)][string]$ApiId,
        [Parameter(Mandatory)][string]$ContractId
    )

    $uri = "$script:AnypointBaseUri/apimanager/api/v1/organizations/$OrgId/environments/$EnvId/apis/$ApiId/contracts/$ContractId/status"
    $body = @{ status = 'APPROVED' }
    return Invoke-AnypointApi -Headers $Headers -Uri $uri -Method PATCH -Body $body
}

function Select-FromList {
    <#
        .SYNOPSIS
        Prints a numbered list built from $Items and returns the item the user picks.
    #>
    param(
        [Parameter(Mandatory)][object[]]$Items,
        [Parameter(Mandatory)][string]$DisplayProperty,
        [Parameter(Mandatory)][string]$Prompt
    )

    if (-not $Items -or $Items.Count -eq 0) {
        throw "No items available to select for: $Prompt"
    }

    Write-Host ""
    Write-Host "== $Prompt ==" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 1), $Items[$i].$DisplayProperty)
    }

    do {
        $selection = Read-Host "Enter the number of your choice"
        $valid = ($selection -as [int]) -and ([int]$selection -ge 1) -and ([int]$selection -le $Items.Count)
        if (-not $valid) { Write-Host "Invalid selection, try again." -ForegroundColor Yellow }
    } while (-not $valid)

    return $Items[[int]$selection - 1]
}

Export-ModuleMember -Function `
    Get-AnypointAccessToken, Get-AnypointAuthHeader, Test-AnypointToken, Invoke-AnypointApi, Get-AnypointOrganizations, Get-AnypointEnvironments, `
    Get-AnypointApiInstances, Get-AnypointContracts, Get-AnypointApplications, Get-AnypointTiers, `
    New-AnypointTier, New-AnypointContract, Approve-AnypointContract, Select-FromList
