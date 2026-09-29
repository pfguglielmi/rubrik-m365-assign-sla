# rubrik-m365-assign-sla

![CI](https://github.com/pfguglielmi/rubrik-m365-assign-sla/actions/workflows/ci.yml/badge.svg)

A PowerShell script for bulk-assigning Rubrik SLA Domains to Microsoft 365 objects in Rubrik Security Cloud (RSC), built on top of Rubrik's [`polaris-o365-powershell`](https://github.com/rubrikinc/polaris-o365-powershell) module.

It supports two targeting modes, for SharePoint sites, OneDrive accounts, or mailboxes (`-ObjectType`):
- Assign an SLA Domain to a **single object** by identifier (`-SearchByUrl`).
- **Bulk-assign** an SLA Domain to a list of objects supplied via a CSV input file (`-InputFile`).

Either mode can also assign the special value `UNPROTECTED`, which instead makes the target object(s) inherit the SLA Domain from their parent. `-WhatIf` previews what a run would do without assigning anything.

## Prerequisites

- PowerShell v6.0 or later
- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation) v1.8.0 or later (`Install-Module Microsoft.Graph`)
- Rubrik's [`polaris-o365-powershell`](https://github.com/rubrikinc/polaris-o365-powershell) module

### 1. Create a custom role in Rubrik Security Cloud

Create a role scoped to Microsoft 365 objects with permission to assign SLA Domains:

1. Log in to RSC → open the app tray → **Settings**.
2. **Users and Access** → **Roles** → **Create Role** → **Custom role**.
3. Name the role, then under **Data Management Permissions** click **Configure**.
4. Choose **By object type** → select **Microsoft 365** and grant the privileges needed (optionally scope to specific objects only).
5. On **Assign SLA Domains**, select the SLA Domain(s) this role should be allowed to assign (or "Select all existing and future SLA Domains").
6. **Create**.

### 2. Create a Service Account

1. RSC → **Settings** → **Users and Access** → **Service Accounts** → **Add Service Account**.
2. Name it, assign it the role created above, and click **Add**.
3. On the final screen, click **Download as JSON** to save the Client ID / Client Secret / Access Token.
4. Rename the downloaded file to `polaris-service-account.json` and place it in a `.rubrik` folder in your home directory (create the folder if it doesn't exist), e.g. `~/.rubrik/polaris-service-account.json`.

> This file contains credentials — never commit it to source control.

### 3. Install the Rubrik module

```powershell
Import-Module ./RubrikPolaris/RubrikPolaris.psd1
```

## Usage

```powershell
# Assign an SLA Domain to a single SharePoint site by URL (SharePoint is the default -ObjectType)
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain '<sla_domain_name>' -SearchByUrl '<site_url>'

# Assign an SLA Domain to a single OneDrive account or mailbox by user principal name / email
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain '<sla_domain_name>' -ObjectType OneDrive -SearchByUrl '<user@tenant.com>'
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain '<sla_domain_name>' -ObjectType Mailbox -SearchByUrl '<user@tenant.com>'

# Bulk-assign an SLA Domain to all objects listed in a CSV file
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain '<sla_domain_name>' -InputFile '<path_to_csv>'

# Either mode: pass 'UNPROTECTED' as -SlaDomain to inherit the SLA Domain from the parent instead
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain 'UNPROTECTED' -SearchByUrl '<site_url>'

# Preview a bulk run without assigning anything
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain '<sla_domain_name>' -InputFile '<path_to_csv>' -WhatIf
```

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-PathToM365Module` | Yes | Path to the Rubrik PowerShell module for M365 protection. |
| `-SubName` | Yes | Name of the Microsoft 365 subscription in RSC to work with. |
| `-SlaDomain` | Yes | Name of the SLA Domain to assign, or `UNPROTECTED` to inherit from the parent. |
| `-SearchByUrl` | One of `-SearchByUrl` / `-InputFile` required (mutually exclusive) | Identifier of a single object to assign the SLA Domain to: a site URL for `-ObjectType SharePoint`, or a user principal name / email for `OneDrive` / `Mailbox`. |
| `-InputFile` | One of `-SearchByUrl` / `-InputFile` required (mutually exclusive) | Path to a semicolon-delimited CSV file with `sitename;URL` columns listing objects to assign the SLA Domain to in bulk. |
| `-ObjectType` | No (default `SharePoint`) | Type of M365 object to target: `SharePoint`, `OneDrive`, or `Mailbox`. See [Supported object types](#supported-object-types). |
| `-WhatIf` | No | Preview the SLA assignments that would be made, without making them. |
| `-Confirm` | No | Prompt for confirmation before each SLA assignment. |

### Supported object types

`-ObjectType` accepts `SharePoint` (default), `OneDrive`, or `Mailbox` — the three Microsoft 365 workloads for which `polaris-o365-powershell` exposes an object-listing cmdlet (`Get-PolarisM365SharePoint`, `Get-PolarisM365OneDrives`, `Get-PolarisM365Mailboxes`). **Teams is not supported**: the module has no `Get-PolarisM365Teams` cmdlet. A Team's file content lives in a SharePoint site, so protecting that Team is done today via `-ObjectType SharePoint` and the Team's site URL (see `sample-input.csv` for examples of Teams-backed sites).

### CSV input format

See [`sample-input.csv`](./sample-input.csv) for the expected format:

```
sitename;URL
Site Display Name;https://tenant.sharepoint.com/sites/SiteName
```

The `URL` column holds the identifier appropriate to `-ObjectType`: a SharePoint site URL for `SharePoint`, or a user principal name / email address for `OneDrive` / `Mailbox`.

## Error handling

- A bad `-SubName`, `-SlaDomain`, or (in single-object mode) an unresolved `-SearchByUrl` target stops the script immediately with a clear error and a non-zero exit code — nothing is assigned.
- In bulk (`-InputFile`) mode, a row whose identifier doesn't match exactly one object, or whose SLA assignment call fails, is skipped with a `Write-Error` message naming the row; the rest of the batch still runs. A summary line at the end reports how many objects were assigned/skipped/declined, and if any row was skipped the script then throws (non-zero exit) so automation still sees the run as failed even though most of the batch succeeded.

## Testing

Pester tests live under [`Tests/`](./Tests) and mock the `polaris-o365-powershell` cmdlets via an in-memory stub module (`Tests/TestHelpers/RubrikPolarisStub`), so they run without a live RSC connection.

```powershell
Install-Module Pester, PSScriptAnalyzer -Scope CurrentUser -Force
Invoke-Pester ./Tests
Invoke-ScriptAnalyzer -Path ./AssignSLA.ps1
```

## CI

[`.github/workflows/ci.yml`](./.github/workflows/ci.yml) runs on every push and pull request: it lints `AssignSLA.ps1` with PSScriptAnalyzer and runs the Pester suite, failing the build on any lint finding or test failure.

## Notes

- The script connects to RSC via `Connect-Polaris`, which reads the service account credentials from `~/.rubrik/polaris-service-account.json`.
- Matching objects by identifier (rather than by name/search string) requires the `url` / `userPrincipalName` properties on objects returned by the module's `Get-PolarisM365*` cmdlets. These were added in relatively recent versions of `polaris-o365-powershell`; if identifier-based matching fails to find *any* objects, check that your copy of the module is recent enough to return them.
- If more than one object of the selected `-ObjectType` shares the same identifier, the script treats that as an error (single-object mode throws; bulk mode skips the row) rather than guessing which one to use.
- Bulk (`-InputFile`) mode reports a `Declined` count in its summary alongside `Assigned`/`Skipped`, covering rows you (or `-WhatIf`) chose not to touch via `-Confirm`.

## License

MIT — see [LICENSE](./LICENSE).
