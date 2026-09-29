@{
    RootModule        = 'RubrikPolarisStub.psm1'
    ModuleVersion     = '1.0'
    GUID              = '5b6f9b8c-3b7d-4c2b-9b8a-1f6a6d3d9a11'
    Author            = 'Test'
    Description       = 'In-memory stand-in for polaris-o365-powershell, used by the Pester suite for AssignSLA.ps1.'
    PowerShellVersion = '6.0'
    FunctionsToExport = @(
        'Connect-Polaris',
        'Get-PolarisM365Subscriptions',
        'Get-PolarisSLA',
        'Get-PolarisM365SharePoint',
        'Get-PolarisM365OneDrives',
        'Get-PolarisM365Mailboxes',
        'Set-PolarisM365ObjectSla',
        'Set-RubrikPolarisStubData',
        'Get-RubrikPolarisStubAssignments',
        'Get-RubrikPolarisStubCalledFunctions'
    )
}
