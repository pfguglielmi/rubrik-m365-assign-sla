# In-memory stand-in for the real polaris-o365-powershell (RubrikPolaris) module.
# AssignSLA.ps1 is unaware this isn't the real module: it's loaded via the same
# -PathToM365Module parameter and exposes the same cmdlet names/signatures. Test
# fixtures are configured via Set-RubrikPolarisStubData before each test.

$script:Subscriptions = @()
$script:Slas = @()
$script:SharePointSites = @()
$script:OneDrives = @()
$script:Mailboxes = @()
$script:Assignments = @()
$script:CalledFunctions = @()
$script:ConnectShouldThrow = $false
$script:AssignShouldThrow = $false

function Set-RubrikPolarisStubData {
    param(
        [array]$Subscriptions = @(),
        [array]$Slas = @(),
        [array]$SharePointSites = @(),
        [array]$OneDrives = @(),
        [array]$Mailboxes = @(),
        [switch]$ConnectShouldThrow,
        [switch]$AssignShouldThrow
    )

    $script:Subscriptions = $Subscriptions
    $script:Slas = $Slas
    $script:SharePointSites = $SharePointSites
    $script:OneDrives = $OneDrives
    $script:Mailboxes = $Mailboxes
    $script:Assignments = @()
    $script:CalledFunctions = @()
    $script:ConnectShouldThrow = [bool]$ConnectShouldThrow
    $script:AssignShouldThrow = [bool]$AssignShouldThrow
}

function Get-RubrikPolarisStubAssignments {
    return $script:Assignments
}

function Get-RubrikPolarisStubCalledFunctions {
    return $script:CalledFunctions
}

function Connect-Polaris {
    if ($script:ConnectShouldThrow) {
        throw 'Stub: unable to connect to Rubrik Security Cloud.'
    }
}

function Get-PolarisM365Subscriptions {
    $script:CalledFunctions += 'Get-PolarisM365Subscriptions'
    return $script:Subscriptions
}

function Get-PolarisSLA {
    param(
        [string]$Name
    )
    $script:CalledFunctions += 'Get-PolarisSLA'
    if ([string]::IsNullOrEmpty($Name)) {
        return $script:Slas
    }
    # Mimic the real (undocumented) server-side filter loosely matching by
    # substring, so the calling script's own exact-name check is exercised.
    return $script:Slas | Where-Object { $_.name -like "*$Name*" }
}

function Get-PolarisM365SharePoint {
    param(
        [string]$SubscriptionId,
        [string]$SearchString,
        [string]$Includes
    )
    $script:CalledFunctions += 'Get-PolarisM365SharePoint'
    return $script:SharePointSites
}

function Get-PolarisM365OneDrives {
    param(
        [string]$SubscriptionId
    )
    $script:CalledFunctions += 'Get-PolarisM365OneDrives'
    return $script:OneDrives
}

function Get-PolarisM365Mailboxes {
    param(
        [string]$SubscriptionId
    )
    $script:CalledFunctions += 'Get-PolarisM365Mailboxes'
    return $script:Mailboxes
}

function Set-PolarisM365ObjectSla {
    param(
        [string[]]$ObjectID,
        [string]$SlaID
    )
    $script:CalledFunctions += 'Set-PolarisM365ObjectSla'
    if ($script:AssignShouldThrow) {
        throw 'Stub: issue assigning SLA domain to object.'
    }
    $script:Assignments += [pscustomobject]@{ ObjectID = $ObjectID; SlaID = $SlaID }
    return 'Success'
}

Export-ModuleMember -Function Connect-Polaris, Get-PolarisM365Subscriptions, Get-PolarisSLA, `
    Get-PolarisM365SharePoint, Get-PolarisM365OneDrives, Get-PolarisM365Mailboxes, Set-PolarisM365ObjectSla, `
    Set-RubrikPolarisStubData, Get-RubrikPolarisStubAssignments, Get-RubrikPolarisStubCalledFunctions
