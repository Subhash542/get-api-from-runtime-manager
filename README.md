# Anypoint API Contract Migration

Copies API **contracts** (client application access grants) from one API
instance in a source **Business Group** to another API instance in a
destination **Business Group**, on Anypoint Platform, via PowerShell.

## What it does

1. Prompts for an Anypoint Platform **bearer token**.
2. Lets you interactively pick:
   - Source Business Group → Environment → API
   - Destination Business Group → Environment → API
3. **Migrates SLA tiers first**: every tier on the source API is matched by
   name against the tiers already on the destination API; any that don't
   exist yet are created there with the same request limits. This gives a
   source-tier-id → destination-tier-id mapping used in the next step.
4. Reads every contract on the source API.
5. For each one:
   - **Already exists on destination** → logged as `Contract already exists`, no change made.
   - **Consumer app found in destination BG, no contract yet** → contract is
     **created and approved**, logged as `Contract created and approved`.
   - **Consumer app not registered in destination BG at all** → skipped and
     logged, since Anypoint contracts require the application to already
     exist in that Business Group.
6. Prints summary tables for both tiers and contracts, and writes two
   timestamped CSVs (`TierMigration_*.csv` and `ContractMigration_*.csv`)
   to `AnypointContractMigrationResults\`.

## Files

```
AnypointContractMigration/
├── Migrate-AnypointApiContracts.ps1   # run this
├── Modules/
│   └── AnypointApi.psm1               # all Anypoint REST calls live here
└── README.md
```

## Prerequisites

- PowerShell 5.1 (Windows PowerShell) or PowerShell 7+.
- Network access to `anypoint.mulesoft.com`.
- A bearer token with permissions on **both** Business Groups involved:
  - View Organization
  - View Environment
  - Manage APIs (or API Manager Contracts management)
  - View/Manage Client Applications
  - Either a **Personal Access Token**, or an **access token from a
    Connected App** with the scopes above.

## Running it

```powershell
cd AnypointContractMigration
.\Migrate-AnypointApiContracts.ps1
```

Optional: pass `-OutputFolder` to control where the CSV report is written:

```powershell
.\Migrate-AnypointApiContracts.ps1 -OutputFolder 'C:\Reports'
```

You'll be prompted for the token (input is masked), then walked through the
Business Group / Environment / API selections for both source and
destination as numbered menus.

## Important assumptions / limitations

- **Applications must already exist in the destination Business Group.**
  An Anypoint "contract" links an *application* to an *API*. If the
  consumer application that holds a contract on the source API has never
  been registered/created in the destination Business Group, this script
  cannot fabricate one — it will log the contract as skipped so you know
  to register that application there first.
- Matching between source and destination applications is done by **Client
  ID** first, falling back to **application name** if Client ID isn't
  available.
- **Revoked** contracts on the source API are not migrated.
- **SLA tiers** are matched/created **by name**. If two tiers on the source
  and destination APIs happen to have the same name but different limits,
  the existing destination tier is reused as-is (its limits are not
  overwritten). If a source contract references a tier that, for any
  reason, doesn't end up in the tier map, the contract is still created but
  without a tier attached, and a warning is printed so you can fix it up
  manually.
- Anypoint contracts don't have a native free-text "comment" field — the
  "already exists" / "created and approved" comments described in this
  project are written to the console output and CSV report, not stored
  back into Anypoint Platform itself.

## Endpoint notes

This script uses the documented Anypoint **Access Management** and
**API Manager v1** REST APIs (`/accounts/api/...`,
`/apimanager/api/v1/...`). All calls are isolated in
`Modules/AnypointApi.psm1`, so if your tenant is on a different API
Manager release with different paths, that's the only file you need to
adjust.
