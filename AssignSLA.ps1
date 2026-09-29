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
    [string]$InputFile
)

$sub = $null
$sla = $null
$csv = $null
$spsites = $null
$site = $null

# Verify that either -SearchByUrl or -InputFile parameter is used
if ([string]::IsNullOrEmpty($SearchByUrl) -and [string]::IsNullOrEmpty($InputFile)) {
    Write-Host "Either the -SearchByUrl or -InputFile parameter must be used to run this script."
    Exit
}
    
# Verify that -SearchByURL and -InputFile are not used simultaneously 
if ($SearchByUrl -and $InputFile) {
    Write-Host "The parameters -SearchByUrl and -InputFile cannot be used simultaneously."
    Exit
}

# Import the PowerShell module for managing Rubrik's Protection for Microsoft 365
Import-Module $PathToM365Module

# Connect to the Rubrik Security Cloud Instance with the PowerShell module for Microsoft 365 protection.. Note that the service account token JSON file must be named polaris-service-account.json and place in the ~/.rubrik folder
Connect-Polaris

# Get the subscription we want to work with
$sub = Get-PolarisM365Subscriptions | Where-Object {$_.name -eq $SubName}

# Get the SLA Domain object we will want to bulk assign to M365 objects$sub
if ($SlaDomain -ne 'UNPROTECTED') {
    $sla = Get-PolarisSLA -Name $SlaDomain
}

# Assign the SLA Domain to the SharePoint site specified by the -SearchByUrl parameter
if ($SearchByUrl) {
    $spsites = Get-PolarisM365SharePoint -SubscriptionId $sub.subscriptionId -Includes 'SitesOnly' | Where-Object url -eq $SearchByUrl
    if ($SlaDomain -eq 'UNPROTECTED') {
        Write-Host "Inheriting SLA Domain from parent and assigning it to site" $spsites.name "which URL is" $spsites.url
        # Assign the SLA Domain inherited from parent to the current SharePoint site
        Set-PolarisM365ObjectSla -ObjectID $spsites.id -SlaID $SlaDomain
    } else {
        Write-Host "Assigning SLA Domain" $SlaDomain "to site" $spsites.name "which URL is" $spsites.url
        # Assign the SLA Domain to the current SharePoint site
        Set-PolarisM365ObjectSla -ObjectID $spsites.id -SlaID $sla.id
    }
}

# Assign the SLA Domain to all SharePoint sites listed in the text file
if ($InputFile) {
    # Import the csv file that contains the list of Microsoft 365 SharePoint sites to protect
    $csv = Import-Csv -Path $InputFile -Delimiter ';' | Select-Object sitename,URL
    $spsites = Get-PolarisM365SharePoint -SubscriptionId $sub.subscriptionId -Includes 'SitesOnly'
    if ($SlaDomain -eq 'UNPROTECTED') {
        foreach ($row in $csv) {
            $site = $spsites | Where-Object url -eq $row.url
            Write-Host "Inheriting SLA Domain from parent and assigning it to site" $site.name "which URL is" $site.url
            # Assign the SLA Domain inherited from parent to the current SharePoint site
            Set-PolarisM365ObjectSla -ObjectID $site.id -SlaID $SlaDomain
        }
    } else {
        foreach ($row in $csv) {
            $site = $spsites | Where-Object url -eq $row.url
            Write-Host "Assigning SLA Domain" $SlaDomain "to site" $site.name "which URL is" $site.url
            # Assign the SLA Domain to the current SharePoint site
            Set-PolarisM365ObjectSla -ObjectID $site.id -SlaID $sla.id
        }
    }
}
