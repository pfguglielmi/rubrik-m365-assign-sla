# rubrik-m365-assign-sla

A PowerShell script and workflow for bulk-assigning Rubrik SLA Domains to Microsoft 365 SharePoint sites in Rubrik Security Cloud (RSC), built on top of Rubrik's [`polaris-o365-powershell`](https://github.com/rubrikinc/polaris-o365-powershell) module.

It supports two modes:
- Assign an SLA Domain to a **single SharePoint site** by URL.
- **Bulk-assign** an SLA Domain to a list of SharePoint sites supplied via a CSV input file.

Either mode can also assign the special value `UNPROTECTED`, which instead makes the target site(s) inherit the SLA Domain from their parent object.

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
# Assign an SLA Domain to a single SharePoint site by URL
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain '<sla_domain_name>' -SearchByUrl '<site_url>'

# Bulk-assign an SLA Domain to all sites listed in a CSV file
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain '<sla_domain_name>' -InputFile '<path_to_csv>'

# Either mode: pass 'UNPROTECTED' as -SlaDomain to inherit the SLA Domain from the parent instead
./AssignSLA.ps1 -PathToM365Module '<path_to_module>' -SubName '<m365_subscription_name>' -SlaDomain 'UNPROTECTED' -SearchByUrl '<site_url>'
```

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-PathToM365Module` | Yes | Path to the Rubrik PowerShell module for M365 protection. |
| `-SubName` | Yes | Name of the Microsoft 365 subscription in RSC to work with. |
| `-SlaDomain` | Yes | Name of the SLA Domain to assign, or `UNPROTECTED` to inherit from the parent. |
| `-SearchByUrl` | One of `-SearchByUrl` / `-InputFile` required (mutually exclusive) | URL of a single SharePoint site to assign the SLA Domain to. |
| `-InputFile` | One of `-SearchByUrl` / `-InputFile` required (mutually exclusive) | Path to a semicolon-delimited CSV file with `sitename;URL` columns listing sites to assign the SLA Domain to in bulk. |

### CSV input format

See [`sample-input.csv`](./sample-input.csv) for the expected format:

```
sitename;URL
Site Display Name;https://tenant.sharepoint.com/sites/SiteName
```

## Notes

- The script connects to RSC via `Connect-Polaris`, which reads the service account credentials from `~/.rubrik/polaris-service-account.json`.
- Matching sites by URL (rather than by name/search string) requires the `url` property on objects returned by `Get-PolarisM365SharePoint`. This property was added upstream in the `polaris-o365-powershell` module.

## License

MIT — see [LICENSE](./LICENSE).
