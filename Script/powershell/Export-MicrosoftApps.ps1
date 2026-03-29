<#
        .SYNOPSIS
        Creates a list of Microsoft first party apps with app id and display name and exports to JSON.

        .DESCRIPTION
        This scripts retrieves a list of apps in the following order
        1. Microsoft Graph (apps where appOwnerOrganizationId is Microsoft)
        2. Microsoft Entra docs (from known-guids.json in the Entra docs repository)
        3. Microsoft Learn doc (https://learn.microsoft.com/troubleshoot/azure/active-directory/verify-first-party-apps-sign-in)
        4. Custom list of apps (./customdata/MysteryApps.csv) - Community contributed list of Microsoft apps and their app ids

        This script connects to Microsoft Graph with the scope Application.Read.All when required.
        .EXAMPLE
        ./Script/powershell/Export-MicrosoftApps.ps1

        Creates a list of Microsoft first party apps with output written to .\data and custom data loaded from ./customdata/OtherMicrosoftApps.csv
        Assumes the root of the repo is the current working directory

        .EXAMPLE
        Export-MicrosoftApps.ps1 -OutputPath ".\myOutputFolder" -CustomAppDataPath ".\customdata\OtherMicrosoftApps.csv"

        Creates a list using custom folders for the output and reading of custom data
#>

param (
    [Parameter(Mandatory=$false, HelpMessage="Path to output the JSON file")]
    [string]$OutputPath = "",

    [Parameter(Mandatory=$false, HelpMessage="Path to csv file with community contributed custom list of apps")]
    [string]$CustomAppDataPath = "",

    [Parameter(Mandatory=$false, HelpMessage="Microsoft Graph access token from Azure CLI or another non-interactive source")]
    [string]$AccessToken
    )

$RepoRoot = if ($PSScriptRoot) {
    Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
} else {
    (Get-Location).Path
}

if (-not $OutputPath) {
    $OutputPath = Join-Path $RepoRoot "data"
}

if (-not $CustomAppDataPath) {
    $CustomAppDataPath = Join-Path $RepoRoot "customdata\\OtherMicrosoftApps.csv"
}

# Ensure connection to Microsoft Graph
try {
    $context = Get-MgContext
    if (-not $context) {
        if ($AccessToken) {
            Write-Host "Connecting to Microsoft Graph with supplied access token..." -ForegroundColor Yellow
            $secureToken = $AccessToken | ConvertTo-SecureString -AsPlainText -Force
            Connect-MgGraph -AccessToken $secureToken -NoWelcome
        } else {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
        }
    } else {
        Write-Host "Already connected to Microsoft Graph as $($context.Account)" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

function GetAppsFromMicrosoftLearnDoc() {
    Write-Debug "Retrieving apps from Microsoft Learn doc"
    $msLearnFirstPartyAppDocUri = "https://raw.githubusercontent.com/MicrosoftDocs/SupportArticles-docs/refs/heads/main/support/entra/entra-id/governance/verify-first-party-apps-sign-in.md"
    $mdContent = (Invoke-WebRequest -Uri $msLearnFirstPartyAppDocUri).Content
    $lines = $mdContent -split "\r?\n"
    $appList = @()
    $inAppTable = $false
    foreach ($line in $lines) {
        $cleanLine = $line.trim()

        if ($cleanLine.startsWith("|")) {
            if ($cleanLine -match '^\|Application Name\|Application IDs\|$') {
                $inAppTable = $true
                continue
            }

            if (-not $inAppTable) { continue }
            if ($cleanLine -match '^\|[-:]+\|[-:]+\|$') { continue }

            $cols = $cleanLine.Trim('|') -split '\|'
            if ($cols.Count -lt 2) { continue }

            $appName = $cols[0].trim()
            $appId = $cols[1].trim().ToLower()

            $guid = [System.Guid]::empty
            $isGuid = [System.Guid]::TryParse($appId, [System.Management.Automation.PSReference]$guid)
            if ($isGuid) {
                $itemInfo = [PSCustomObject]@{
                    AppId                  = $appId + ""
                    AppDisplayName         = $appName + ""
                    AppOwnerOrganizationId = "72f988bf-86f1-41af-91ab-2d7cd011db47"
                    Source                 = "Learn"
                }
                $appList += $itemInfo
            }
            continue
        }

        if ($inAppTable -and -not [string]::IsNullOrWhiteSpace($cleanLine)) {
            break
        }
    }
    Write-Host "  Found $($appList.Count) apps from Microsoft Learn" -ForegroundColor Green
    return $appList
}

function GetAppsFromEntraDocs() {
    Write-Host "Retrieving apps from Entra documentation source"
    $docsJsonUri = "https://raw.githubusercontent.com/MicrosoftDocs/entra-docs/main/.docutune/dictionaries/known-guids.json"

    try {
        $response = Invoke-WebRequest -Uri $docsJsonUri -ErrorAction Stop
        $rawContent = $response.Content
        
        $appList = @()
        $seenGuids = @{}
        
        # Parse JSON manually using regex to handle duplicate keys and comments
        # Pattern matches: "App Name" : "GUID"
        $pattern = '"([^"]+)"\s*:\s*"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"'
        
        $matches = [regex]::Matches($rawContent, $pattern)
        
        foreach ($match in $matches) {
            $appDisplayName = $match.Groups[1].Value
            $appId = $match.Groups[2].Value.ToLower()
            
            # Skip duplicates (keep first occurrence)
            if ($seenGuids.ContainsKey($appId)) { continue }
            $seenGuids[$appId] = $true
            
            $itemInfo = [PSCustomObject]@{
                AppId                  = $appId
                AppDisplayName         = $appDisplayName
                AppOwnerOrganizationId = ""
                Source                 = "EntraDocs"
            }
            $appList += $itemInfo
        }

        Write-Host "  Found $($appList.Count) apps from Entra documentation"
        return $appList
    }
    catch {
        Write-Error "Failed to retrieve data from Entra documentation: $_"
        return @()
    }
}

function GetAppsFromMicrosoftGraph() {
    Write-Host "Retrieving apps from Microsoft Graph" -ForegroundColor Cyan
    $tenantIdList = @("f8cdef31-a31e-4b4a-93e4-5f571e91255a", "72f988bf-86f1-41af-91ab-2d7cd011db47", "cdc5aeea-15c5-4db6-b079-fcadd2505dc2")
    $select = "appId,appDisplayName,appOwnerOrganizationId"
    $servicePrincipals = @()

    try {
        foreach ($tenantId in $tenantIdList) {
            $filter = "appOwnerOrganizationId eq $($tenantId)"
            $servicePrincipals += Get-MgServicePrincipal -Filter $filter -Select $select -ConsistencyLevel eventual -PageSize 999 -CountVariable $count -All
        }
    } catch {
        Write-Error "Failed to retrieve Microsoft Graph service principals: $_"
        return @()
    }

    $appList = @()

    foreach ($item in $servicePrincipals) {
        $itemInfo = [PSCustomObject]@{
            AppId                  = $item.appId + ""
            AppDisplayName         = $item.appDisplayName + ""
            AppOwnerOrganizationId = $item.appOwnerOrganizationId + ""
            Source                 = "Graph"
        }
        $appList += $itemInfo
    }

    Write-Host "  Found $($appList.Count) apps from Microsoft Graph" -ForegroundColor Green
    return $appList
}

$appList = @()

# Sources at the top take priority, duplicates from sources that are lower are skipped.
$appList += GetAppsFromMicrosoftGraph
$appList += GetAppsFromEntraDocs
$appList += GetAppsFromMicrosoftLearnDoc

if (Test-Path $CustomAppDataPath) {
    $customApps = Import-Csv $CustomAppDataPath | Where-Object {
        $_.AppId -and $_.AppDisplayName
    } | ForEach-Object {
        $_.AppDisplayName = $_.AppDisplayName.Trim() + " [Community Contributed]"
        $_
    }

    Write-Host "Loaded $($customApps.Count) community-contributed apps" -ForegroundColor Green
    $appList += $customApps
} else {
    Write-Warning "Custom app data file not found at $CustomAppDataPath. Continuing without community data."
}

Write-Host "Creating unique list of apps"
$appMap = @{}

foreach ($item in $appList) {
    [string]$id = ($item.AppId + "").ToLower()
    if (-not $id) { continue }

    if (-not $appMap.ContainsKey($id)) {
        $appMap[$id] = [ordered]@{
            AppId                  = $id
            AppDisplayName         = $item.AppDisplayName + ""
            AppOwnerOrganizationId = $item.AppOwnerOrganizationId + ""
            Source                 = $item.Source + ""
            Sources                = @($item.Source + "")
        }
        continue
    }

    $existing = $appMap[$id]
    if (-not [string]::IsNullOrWhiteSpace($item.AppOwnerOrganizationId) -and [string]::IsNullOrWhiteSpace($existing.AppOwnerOrganizationId)) {
        $existing.AppOwnerOrganizationId = $item.AppOwnerOrganizationId + ""
    }
    if ([string]::IsNullOrWhiteSpace($existing.AppDisplayName) -and -not [string]::IsNullOrWhiteSpace($item.AppDisplayName)) {
        $existing.AppDisplayName = $item.AppDisplayName + ""
    }
    if ($existing.Sources -notcontains ($item.Source + "")) {
        $existing.Sources += ($item.Source + "")
    }
}

$uniqueAppList = $appMap.GetEnumerator() | Sort-Object { $_.Value.AppDisplayName } | ForEach-Object {
    [PSCustomObject]$_.Value
}

Write-Host "Exporting to JSON"
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

$outputFilePathJson = Join-Path $OutputPath "MicrosoftApps.json"

@($uniqueAppList) | ConvertTo-Json -Depth 10 | Out-File $outputFilePathJson -Encoding UTF8

Write-Host "Export complete:" -ForegroundColor Green
($uniqueAppList | Group-Object Source | Sort-Object Name) | ForEach-Object {
    Write-Host "  - $($_.Name): $($_.Count)" -ForegroundColor Green
}
Write-Host "  - Total unique apps: $($uniqueAppList.Count)" -ForegroundColor Green
