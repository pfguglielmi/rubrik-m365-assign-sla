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
$slaIdToAssign = if ($SlaDomain -eq 'UNPROTECTED') { $SlaDomain } else { $sla.id }

# Assign the SLA Domain to the single object specified by the -SearchByUrl parameter
# (the site URL for -ObjectType SharePoint, or the user's email/UPN for OneDrive/Mailbox)
if ($SearchByUrl) {
    $inventory = Get-M365ObjectInventory -ObjectType $ObjectType -SubscriptionId $sub.subscriptionId
    $target = $inventory | Where-Object Identifier -eq $SearchByUrl

    if (-not $target) {
        throw "No $ObjectType object matching '$SearchByUrl' was found in subscription '$SubName'."
    }

    if ($PSCmdlet.ShouldProcess($target.Name, "Assign SLA Domain '$SlaDomain'")) {
        if ($SlaDomain -eq 'UNPROTECTED') {
            Write-Host "Inheriting SLA Domain from parent and assigning it to $ObjectType object" $target.Name "which identifier is" $target.Identifier
        } else {
            Write-Host "Assigning SLA Domain" $SlaDomain "to $ObjectType object" $target.Name "which identifier is" $target.Identifier
        }
        Set-PolarisM365ObjectSla -ObjectID $target.Id -SlaID $slaIdToAssign
    }
}

# Assign the SLA Domain to all objects listed in the input file
if ($InputFile) {
    # Import the csv file that contains the list of Microsoft 365 objects to protect
    $csv = Import-Csv -Path $InputFile -Delimiter ';' | Select-Object sitename, URL
    $inventory = Get-M365ObjectInventory -ObjectType $ObjectType -SubscriptionId $sub.subscriptionId

    $assignedCount = 0
    $skippedCount = 0

    foreach ($row in $csv) {
        $target = $inventory | Where-Object Identifier -eq $row.url

        if (-not $target) {
            Write-Error "No $ObjectType object matching '$($row.url)' (row '$($row.sitename)') was found; skipping."
            $skippedCount++
            continue
        }

        if ($PSCmdlet.ShouldProcess($target.Name, "Assign SLA Domain '$SlaDomain'")) {
            if ($SlaDomain -eq 'UNPROTECTED') {
                Write-Host "Inheriting SLA Domain from parent and assigning it to $ObjectType object" $target.Name "which identifier is" $target.Identifier
            } else {
                Write-Host "Assigning SLA Domain" $SlaDomain "to $ObjectType object" $target.Name "which identifier is" $target.Identifier
            }
            try {
                Set-PolarisM365ObjectSla -ObjectID $target.Id -SlaID $slaIdToAssign -ErrorAction Stop
                $assignedCount++
            }
            catch {
                Write-Error "Failed to assign SLA Domain to $($target.Name): $_"
                $skippedCount++
            }
        }
    }

    Write-Host "Done. Assigned: $assignedCount. Skipped/failed: $skippedCount."
}
