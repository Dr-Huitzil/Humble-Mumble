<#
.SYNOPSIS
    Export_AD_Computers_Raw.ps1
    Exports a raw CSV of all Active Directory computer objects containing
    the exact selected columns, 3-part OU split, and elapsed day calculations.

.DESCRIPTION
    Extracts:
    - Name
    - Status (Enabled/Disabled)
    - Column1 (1st OU/container token after removing CN)
    - Column2 (2nd OU/container token)
    - Column3 (3rd OU/container token)
    - OperatingSystem
    - OperatingSystemVersion
    - LastLogonDate
    - Last Logon - Days
    - CreatedDate
    - Created - Days
    - ModifiedDate
    - Modified- Days
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ".\Computers_Raw_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Ensure ActiveDirectory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "The 'ActiveDirectory' PowerShell module was not found. Please run this on a Domain Controller or a workstation with RSAT installed."
    exit 1
}

Import-Module ActiveDirectory -ErrorAction Stop

Write-Host "Querying Active Directory for all computer objects..." -ForegroundColor Cyan

$properties = @(
    "Name",
    "Enabled",
    "OperatingSystem",
    "OperatingSystemVersion",
    "LastLogonDate",
    "whenCreated",
    "whenChanged",
    "DistinguishedName"
)

$computers = Get-ADComputer -Filter * -Properties $properties
$now = Get-Date

Write-Host "Found $($computers.Count) computer objects. Generating streamlined raw CSV..." -ForegroundColor Green

$results = foreach ($c in $computers) {
    # Split DistinguishedName by comma, drop the first element (CN=<Name>), and keep the next 3 tokens
    $dnParts = $c.DistinguishedName -split ','
    $ouParts = if ($dnParts.Count -gt 1) { $dnParts[1..($dnParts.Count - 1)] } else { @() }
    
    $col1 = if ($ouParts.Count -gt 0) { $ouParts[0].Trim() } else { "" }
    $col2 = if ($ouParts.Count -gt 1) { $ouParts[1].Trim() } else { "" }
    $col3 = if ($ouParts.Count -gt 2) { $ouParts[2].Trim() } else { "" }

    # Calculate days elapsed
    $lastLogonDays = if ($c.LastLogonDate) { [int][math]::Floor(($now - $c.LastLogonDate).TotalDays) } else { "" }
    $createdDays   = if ($c.whenCreated)   { [int][math]::Floor(($now - $c.whenCreated).TotalDays) } else { "" }
    $modifiedDays  = if ($c.whenChanged)   { [int][math]::Floor(($now - $c.whenChanged).TotalDays) } else { "" }

    [PSCustomObject]@{
        "Name"                   = $c.Name
        "Status"                 = if ($c.Enabled) { "Enabled" } else { "Disabled" }
        "Column1"                = $col1
        "Column2"                = $col2
        "Column3"                = $col3
        "OperatingSystem"        = $c.OperatingSystem
        "OperatingSystemVersion" = $c.OperatingSystemVersion
        "LastLogonDate"          = if ($c.LastLogonDate) { $c.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        "Last Logon - Days"      = $lastLogonDays
        "CreatedDate"            = if ($c.whenCreated)   { $c.whenCreated.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        "Created - Days"         = $createdDays
        "ModifiedDate"           = if ($c.whenChanged)   { $c.whenChanged.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        "Modified- Days"         = $modifiedDays
    }
}

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Export completed successfully!" -ForegroundColor Green
Write-Host "Saved to: $(Resolve-Path $OutputPath)" -ForegroundColor Yellow
