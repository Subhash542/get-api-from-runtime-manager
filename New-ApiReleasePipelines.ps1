param(
  [string]$Org        = "yourorg",
  [string]$Project    = "yourproject",
  [int]   $TemplateId = 12,
  [string]$CsvPath    = ".\apis.csv",
  [string]$NamePrefix = "Release-",
  [string]$Branch     = "refs/heads/main",
  [switch]$DisableCDTrigger
)
$ErrorActionPreference = 'Stop'

$pat    = $env:AZDO_PAT
$hdr    = @{ Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat")) }
$relApi = "https://vsrm.dev.azure.com/$Org/$Project/_apis/release"
$bldApi = "https://dev.azure.com/$Org/$Project/_apis/build/definitions"
$v      = "api-version=7.1"

function Set-Var($vars, $name, $value) {
  if ($vars.PSObject.Properties[$name]) { $vars.$name.value = $value }
  else { $vars | Add-Member -NotePropertyName $name -NotePropertyValue ([pscustomobject]@{ value = $value }) }
}

# Template and existing pipeline names
$template = Invoke-RestMethod "$relApi/definitions/$TemplateId`?$v" -Headers $hdr
$existing = (Invoke-RestMethod "$relApi/definitions?$v" -Headers $hdr).value.name

$results = foreach ($row in (Import-Csv $CsvPath)) {
  $name = "$NamePrefix$($row.ApiName)"
  try {
    if ($existing -contains $name) {
      Write-Host "SKIP $name (already exists)"
      [pscustomobject]@{ Api = $row.ApiName; Status = "Skipped" }
      continue
    }

    # This API's build pipeline (exact name match)
    $build = (Invoke-RestMethod "$bldApi`?name=$([uri]::EscapeDataString($row.BuildName))&$v" -Headers $hdr).value |
             Where-Object name -eq $row.BuildName | Select-Object -First 1
    if (-not $build) { throw "Build pipeline '$($row.BuildName)' not found" }

    # Clone template, strip server-generated fields
    $def = $template | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    foreach ($p in 'id','revision','url','_links','createdBy','createdOn','modifiedBy','modifiedOn') {
      $def.PSObject.Properties.Remove($p)
    }
    $def.name = $name

    # Point artifact at this API's build pipeline (alias unchanged)
    $art = $def.artifacts | Select-Object -First 1
    $art.sourceId = "$($build.project.id):$($build.id)"
    $ref = $art.definitionReference
    $ref.definition.id   = "$($build.id)"
    $ref.definition.name = $build.name
    if ($ref.project) {
      $ref.project.id   = $build.project.id
      $ref.project.name = $build.project.name
    }

    # Default version: latest build from the specified branch
    foreach ($p in 'defaultVersion','defaultVersionSpecific','defaultVersionTags','defaultVersionBranch','defaultVersionType') {
      if ($ref.PSObject.Properties[$p]) { $ref.PSObject.Properties.Remove($p) }
    }
    $ref | Add-Member -NotePropertyName defaultVersionType   -NotePropertyValue ([pscustomobject]@{ id = "latestFromBranchType"; name = "Latest from a specific branch with tags" })
    $ref | Add-Member -NotePropertyName defaultVersionBranch -NotePropertyValue ([pscustomobject]@{ id = $Branch; name = $Branch })

    # Continuous deployment trigger: disable, or restrict to the branch
    if ($DisableCDTrigger) {
      $def.triggers = @()
    } else {
      foreach ($t in $def.triggers) {
        if ($t.triggerType -eq 'artifactSource') {
          $t.triggerConditions = @([pscustomobject]@{
            sourceBranch = $Branch; tags = @()
            useBuildDefinitionBranch = $false; createReleaseOnBuildTagging = $false })
        }
      }
    }

    # Per-API variable
    Set-Var $def.variables 'ApiName' $row.ApiName

    $created = Invoke-RestMethod "$relApi/definitions?$v" -Method Post -Headers $hdr `
      -ContentType "application/json" -Body ($def | ConvertTo-Json -Depth 100)
    Write-Host "OK   $name -> id $($created.id) (build: $($build.name), branch: $Branch)"
    [pscustomobject]@{ Api = $row.ApiName; Status = "Created"; DefinitionId = $created.id; Build = $build.name }
  }
  catch {
    Write-Warning "FAIL $name : $($_.Exception.Message)"
    [pscustomobject]@{ Api = $row.ApiName; Status = "Failed"; Error = $_.Exception.Message }
  }
}

$results | Format-Table -AutoSize
$results | Export-Csv .\pipeline-creation-results.csv -NoTypeInformation
