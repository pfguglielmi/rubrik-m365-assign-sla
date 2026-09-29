[CmdletBinding(SupportsShouldProcess)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive CLI script; progress messages are intentionally always visible, unlike Write-Information/-Verbose.')]
param(
    [Parameter(Mandatory)]
    [string]$PathToM365Module,
    [Parameter(Mandatory)]
    [string]$SubName,
    [Parameter(Mandatory)]
    [string]$SlaDomain,
    [Parameter(Mandatory=$false)]
    [string]$SearchByUrl,
    [Parameter(Mandatory=$false)]
    [string]$InputFile,
    [Parameter(Mandatory=$false)]
    [ValidateSet('SharePoint', 'OneDrive', 'Mailbox')]
    [string]$ObjectType = 'SharePoint'
)

# Returns every object of -ObjectType in the subscription, normalized to Name/Identifier/Id,
# so the rest of the script can resolve a target without caring which underlying cmdlet was used.
function Get-M365ObjectInventory {
    param(
        [Parameter(Mandatory)]
        [string]$ObjectType,
        [Parameter(Mandatory)]
        [string]$SubscriptionId
    )

    switch ($ObjectType) {
        'SharePoint' {
            Get-PolarisM365SharePoint -SubscriptionId $SubscriptionId -Includes 'SitesOnly' |
                Select-Object @{Name = 'Name'; Expression = { $_.name } },
                              @{Name = 'Identifier'; Expression = { $_.url } },
                              @{Name = 'Id'; Expression = { $_.id } }
        }
        'OneDrive' {
            Get-PolarisM365OneDrives -SubscriptionId $SubscriptionId |
                Select-Object @{Name = 'Name'; Expression = { $_.name } },
                              @{Name = 'Identifier'; Expression = { $_.userPrincipalName } },
                              @{Name = 'Id'; Expression = { $_.id } }
        }
        'Mailbox' {
            Get-PolarisM365Mailboxes -SubscriptionId $SubscriptionId |
                Select-Object @{Name = 'Name'; Expression = { $_.name } },
                              @{Name = 'Identifier'; Expression = { $_.userPrincipalName } },
                              @{Name = 'Id'; Expression = { $_.id } }
        }
        default {
            throw "Unsupported -ObjectType '$ObjectType'."
        }
    }
}

# Builds a lookup from Identifier to the (possibly more than one) inventory objects sharing it,
# so callers can do an O(1) lookup per target instead of an O(n) scan, and detect ambiguous matches.
function Get-M365ObjectIndex {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Inventory
    )

    $index = @{}
    foreach ($obj in $Inventory) {
        if (-not $index.ContainsKey($obj.Identifier)) {
            $index[$obj.Identifier] = @()
        }
        $index[$obj.Identifier] += $obj
    }
    return $index
}

function Resolve-M365Target {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Index,
        [Parameter(Mandatory)]
        [string]$Identifier,
        [Parameter(Mandatory)]
        [string]$ObjectType
    )

    if (-not $Index.ContainsKey($Identifier)) {
        return [pscustomobject]@{ Target = $null; Error = "No $ObjectType object matching '$Identifier' was found." }
    }

    $candidates = @($Index[$Identifier])
    if ($candidates.Count -gt 1) {
        return [pscustomobject]@{ Target = $null; Error = "Multiple $ObjectType objects matching '$Identifier' were found; unable to determine which one to use." }
    }

    return [pscustomobject]@{ Target = $candidates[0]; Error = $null }
}

# Performs (or, under -WhatIf/-Confirm, previews) a single SLA assignment and reports what happened,
# so both the single-object and bulk code paths share identical ShouldProcess/error-handling behavior.
# Takes the calling script's own $PSCmdlet (via -Cmdlet) so ShouldProcess reflects its -WhatIf/-Confirm
# state; SupportsShouldProcess is declared here too only to satisfy PSScriptAnalyzer's ShouldProcess rule.
function Invoke-M365SlaAssignment {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        $Target,
        [Parameter(Mandatory)]
        [string]$ObjectType,
        [Parameter(Mandatory)]
        [string]$SlaDomain,
        [Parameter(Mandatory)]
        [string]$SlaIdToAssign,
        [Parameter(Mandatory)]
        $Cmdlet
    )

    if (-not $Cmdlet.ShouldProcess($Target.Name, "Assign SLA Domain '$SlaDomain'")) {
        return 'Declined'
    }

    if ($SlaDomain -eq 'UNPROTECTED') {
        Write-Host "Inheriting SLA Domain from parent and assigning it to $ObjectType object" $Target.Name "which identifier is" $Target.Identifier
    } else {
        Write-Host "Assigning SLA Domain" $SlaDomain "to $ObjectType object" $Target.Name "which identifier is" $Target.Identifier
    }

    try {
        Set-PolarisM365ObjectSla -ObjectID $Target.Id -SlaID $SlaIdToAssign -ErrorAction Stop
        return 'Assigned'
    }
    catch {
        # -ErrorAction Continue guarantees this stays non-terminating (so the caller's loop can
        # keep going) regardless of the caller's ambient $ErrorActionPreference -- notably, CI
        # runners commonly default pwsh steps to $ErrorActionPreference = 'Stop'.
        Write-Error "Failed to assign SLA Domain to $($Target.Name): $_" -ErrorAction Continue
        return 'Failed'
    }
}

# Verify that either -SearchByUrl or -InputFile parameter is used
if ([string]::IsNullOrEmpty($SearchByUrl) -and [string]::IsNullOrEmpty($InputFile)) {
    throw "Either the -SearchByUrl or -InputFile parameter must be used to run this script."
}

# Verify that -SearchByURL and -InputFile are not used simultaneously
if ($SearchByUrl -and $InputFile) {
    throw "The parameters -SearchByUrl and -InputFile cannot be used simultaneously."
}

# In bulk mode, fail fast if the CSV doesn't exist rather than letting Import-Csv error out later
if ($InputFile -and -not (Test-Path -Path $InputFile -PathType Leaf)) {
    throw "The input file '$InputFile' was not found."
}

# Import the PowerShell module for managing Rubrik's Protection for Microsoft 365
try {
    Import-Module $PathToM365Module -ErrorAction Stop
}
catch {
    throw "Failed to import the M365 module from '$PathToM365Module': $_"
}

# Connect to the Rubrik Security Cloud Instance with the PowerShell module for Microsoft 365 protection.. Note that the service account token JSON file must be named polaris-service-account.json and place in the ~/.rubrik folder
try {
    Connect-Polaris -ErrorAction Stop
}
catch {
    throw "Failed to connect to Rubrik Security Cloud: $_"
}

# Get the subscription we want to work with
$matchingSubs = @(Get-PolarisM365Subscriptions | Where-Object { $_.name -eq $SubName })
if ($matchingSubs.Count -eq 0) {
    throw "No Microsoft 365 subscription named '$SubName' was found."
}
if ($matchingSubs.Count -gt 1) {
    throw "Multiple Microsoft 365 subscriptions named '$SubName' were found; unable to determine which one to use."
}
$sub = $matchingSubs[0]

# Get the SLA Domain object we will want to bulk assign to M365 objects
$sla = $null
if ($SlaDomain -ne 'UNPROTECTED') {
    $matchingSlas = @(Get-PolarisSLA -Name $SlaDomain | Where-Object { $_.name -eq $SlaDomain })
    if ($matchingSlas.Count -eq 0) {
        throw "No SLA Domain named '$SlaDomain' was found."
    }
    if ($matchingSlas.Count -gt 1) {
        throw "Multiple SLA Domains named '$SlaDomain' were found; unable to determine which one to use."
    }
    $sla = $matchingSlas[0]
}
# Forward the canonical sentinel (rather than the user's original casing) since -eq/-ne above
# already matched it case-insensitively, but the value we hand to Set-PolarisM365ObjectSla should
# be unambiguous regardless of how that cmdlet implements its own sentinel comparison.
$slaIdToAssign = if ($SlaDomain -eq 'UNPROTECTED') { 'UNPROTECTED' } else { $sla.id }

$inventoryIndex = Get-M365ObjectIndex -Inventory @(Get-M365ObjectInventory -ObjectType $ObjectType -SubscriptionId $sub.subscriptionId)

# Assign the SLA Domain to the single object specified by the -SearchByUrl parameter
# (the site URL for -ObjectType SharePoint, or the user's email/UPN for OneDrive/Mailbox)
if ($SearchByUrl) {
    $resolved = Resolve-M365Target -Index $inventoryIndex -Identifier $SearchByUrl -ObjectType $ObjectType
    if (-not $resolved.Target) {
        throw $resolved.Error
    }

    $status = Invoke-M365SlaAssignment -Target $resolved.Target -ObjectType $ObjectType -SlaDomain $SlaDomain `
        -SlaIdToAssign $slaIdToAssign -Cmdlet $PSCmdlet
    if ($status -eq 'Failed') {
        throw "Failed to assign SLA Domain '$SlaDomain' to $($resolved.Target.Name)."
    }
}

# Assign the SLA Domain to all objects listed in the input file
if ($InputFile) {
    # Import the csv file that contains the list of Microsoft 365 objects to protect
    $csv = Import-Csv -Path $InputFile -Delimiter ';' | Select-Object sitename, URL

    $assignedCount = 0
    $declinedCount = 0
    $skippedCount = 0

    foreach ($row in $csv) {
        $resolved = Resolve-M365Target -Index $inventoryIndex -Identifier $row.url -ObjectType $ObjectType
        if (-not $resolved.Target) {
            # -ErrorAction Continue: see the comment in Invoke-M365SlaAssignment's catch block.
            Write-Error "$($resolved.Error) (row '$($row.sitename)'); skipping." -ErrorAction Continue
            $skippedCount++
            continue
        }

        $status = Invoke-M365SlaAssignment -Target $resolved.Target -ObjectType $ObjectType -SlaDomain $SlaDomain `
            -SlaIdToAssign $slaIdToAssign -Cmdlet $PSCmdlet
        switch ($status) {
            'Assigned' { $assignedCount++ }
            'Declined' { $declinedCount++ }
            'Failed' { $skippedCount++ }
        }
    }

    Write-Host "Done. Assigned: $assignedCount. Skipped/failed: $skippedCount. Declined: $declinedCount."

    if ($skippedCount -gt 0) {
        throw "$skippedCount of $($csv.Count) objects in '$InputFile' could not be assigned the SLA Domain '$SlaDomain'. See the errors above."
    }
}
