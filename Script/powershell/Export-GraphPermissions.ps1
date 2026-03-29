<#
    .SYNOPSIS
    Creates a list of Microsoft Graph permission roles.

    .DESCRIPTION
    Exports Application (AppRoles) and Delegated (Oauth2PermissionScopes) permissions
    from the Microsoft Graph service principal with full details.

    .EXAMPLE
    ./Script/powershell/Export-GraphPermissions.ps1

    Creates a list of all Graph permissions with output written to .\data

    .EXAMPLE
    Export-GraphPermissions.ps1 -OutputPath ".\myOutputFolder"

    Creates a list using custom folders for the output
#>

param (
    [Parameter(Mandatory=$false, HelpMessage="Path to output the JSON files")]
    [string]$OutputPath = "",

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

function GetPermissionsFromMicrosoftGraph() {
    Write-Host "Retrieving permissions from Microsoft Graph..." -ForegroundColor Cyan
    $graphAppId = "00000003-0000-0000-c000-000000000000"

    try {
        $sp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -All
        Write-Host "  Found $($sp.AppRoles.Count) Application permissions" -ForegroundColor Green
        Write-Host "  Found $($sp.Oauth2PermissionScopes.Count) Delegated permissions" -ForegroundColor Green
        return $sp
    } catch {
        Write-Error "Failed to retrieve Graph service principal: $_"
        return $null
    }
}

$sp = GetPermissionsFromMicrosoftGraph

if (-not $sp) {
    Write-Error "Could not retrieve permissions. Exiting."
    exit 1
}

Write-Host "`nExporting to JSON..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

# Export Application Roles
$outputFilePathJson = Join-Path $OutputPath "GraphAppRoles.json"

$appRoles = $sp.AppRoles | Select-Object Id, Value, DisplayName, Description, IsEnabled, Origin
$appRoles | ConvertTo-Json -Depth 10 | Out-File $outputFilePathJson -Encoding UTF8

Write-Host "  - Exported $($appRoles.Count) Application permissions" -ForegroundColor Green

# Export Delegated Permissions
$outputFilePathJson = Join-Path $OutputPath "GraphDelegateRoles.json"

$delegateRoles = $sp.Oauth2PermissionScopes | Select-Object Id, Value, AdminConsentDisplayName, AdminConsentDescription, UserConsentDisplayName, UserConsentDescription, Type, IsEnabled
$delegateRoles | ConvertTo-Json -Depth 10 | Out-File $outputFilePathJson -Encoding UTF8

Write-Host "  - Exported $($delegateRoles.Count) Delegated permissions" -ForegroundColor Green

Write-Host "Export complete!" -ForegroundColor Green
