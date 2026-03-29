<#
.SYNOPSIS
  Graph Permission "Bulldozer" Parser
  Author: Cengiz YILMAZ (MVP)
  Description: Ignores YAML parsing rules, performs direct text mining.
  Parses both v1.0 and beta Microsoft Graph OpenAPI specifications.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

function Get-TempDirectory {
    $candidates = @(
        $env:RUNNER_TEMP,
        $env:TEMP,
        $env:TMPDIR,
        $env:TMP
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($candidate in $candidates) {
        if (-not (Test-Path $candidate)) {
            New-Item -ItemType Directory -Path $candidate -Force | Out-Null
        }

        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $fallback = [System.IO.Path]::GetTempPath()
    if ([string]::IsNullOrWhiteSpace($fallback)) {
        throw "No temporary directory is available."
    }

    if (-not (Test-Path $fallback)) {
        New-Item -ItemType Directory -Path $fallback -Force | Out-Null
    }

    return $fallback
}

function Get-OpenApiCacheDirectory {
    $preferred = if (-not [string]::IsNullOrWhiteSpace($env:GRAPH_OPENAPI_CACHE_DIR)) {
        $env:GRAPH_OPENAPI_CACHE_DIR
    } else {
        Join-Path -Path (Get-TempDirectory) -ChildPath "graph-openapi-cache"
    }

    if (-not (Test-Path $preferred)) {
        New-Item -ItemType Directory -Path $preferred -Force | Out-Null
    }

    return $preferred
}

# --- PATH CONFIGURATION ---
if ($PSScriptRoot) { 
    $BasePath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
} else { 
    $BasePath = "C:\GraphPermission" 
}

$DataPath = if ($OutputPath) { $OutputPath } else { Join-Path -Path $BasePath -ChildPath "data" }
$OutCsv = Join-Path -Path $DataPath -ChildPath "permission.csv"

$CacheDirectory = Get-OpenApiCacheDirectory

# API Sources
$ApiSources = @(
    @{
        Name = "v1.0"
        Url = "https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/v1.0/openapi.yaml"
        TempFile = Join-Path -Path $CacheDirectory -ChildPath "openapi_graph_v1.yaml"
        MetaFile = Join-Path -Path $CacheDirectory -ChildPath "openapi_graph_v1.meta.json"
    },
    @{
        Name = "beta"
        Url = "https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/beta/openapi.yaml"
        TempFile = Join-Path -Path $CacheDirectory -ChildPath "openapi_graph_beta.yaml"
        MetaFile = Join-Path -Path $CacheDirectory -ChildPath "openapi_graph_beta.meta.json"
    }
)

Write-Host "Working Path: $BasePath" -ForegroundColor Gray
Write-Host "Data Path: $DataPath" -ForegroundColor Gray

# Ensure data directory exists
if (-not (Test-Path $DataPath)) {
    New-Item -ItemType Directory -Path $DataPath -Force | Out-Null
}

$AllResults = [System.Collections.Generic.List[PSCustomObject]]::new()

# Process each API version
foreach ($Api in $ApiSources) {
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "Processing $($Api.Name) API..." -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan

    $cacheValid = $false
    if ((Test-Path $Api.TempFile) -and (Test-Path $Api.MetaFile)) {
        try {
            $meta = Get-Content $Api.MetaFile | ConvertFrom-Json
            $cacheAge = (Get-Date) - [DateTime]$meta.downloaded
            if ($cacheAge.TotalHours -lt 24) {
                $cacheValid = $true
                Write-Host "Using cached OpenAPI specification ($($Api.Name), age: $([math]::Round($cacheAge.TotalHours, 1))h)..." -ForegroundColor Yellow
            }
        } catch {
            $cacheValid = $false
        }
    }

    if (-not $cacheValid) {
        Write-Host "Downloading OpenAPI specification ($($Api.Name))..." -ForegroundColor Yellow
        try {
            Invoke-WebRequest -Uri $Api.Url -OutFile $Api.TempFile -UseBasicParsing
            @{ downloaded = (Get-Date).ToString("o"); url = $Api.Url } | ConvertTo-Json | Set-Content $Api.MetaFile
            Write-Host "Download completed." -ForegroundColor Green
        } catch {
            Write-Warning "WARNING: Failed to download $($Api.Name) OpenAPI file. $_"
            if (-not (Test-Path $Api.TempFile)) {
                continue
            }
            Write-Host "Falling back to cached OpenAPI specification ($($Api.Name))." -ForegroundColor Yellow
        }
    }

    $FileSize = (Get-Item $Api.TempFile).Length
    Write-Host "File size: $([math]::Round($FileSize / 1MB, 2)) MB" -ForegroundColor Gray
    
    if ($FileSize -lt 1000000) { # Less than 1MB is suspicious
        Write-Warning "WARNING: File size is very small ($($FileSize / 1KB) KB). It might be an HTML error page."
        Write-Host "First 5 lines of the file:" -ForegroundColor Cyan
        Get-Content $Api.TempFile -TotalCount 5
        Write-Host "---"
        continue
    }

    # 2. Bulldozer Mode: Line by Line Reading
    Write-Host "Scanning $($Api.Name) file..." -ForegroundColor Cyan

    $Reader = [System.IO.File]::OpenText($Api.TempFile)

    # State Variables
    $CurrentPath = "Unknown"
    $CurrentMethod = "Unknown"
    $Counter = 0

    while (($Line = $Reader.ReadLine()) -ne $null) {
        $Counter++
        if ($Counter % 100000 -eq 0) { Write-Host "   -> $Counter lines processed..." -ForegroundColor DarkGray }

        $Trimmed = $Line.Trim()

        # A) PATH DETECTION: Line starts with "/" and ends with ":"
        if ($Trimmed -match "^['""]?(/[a-zA-Z0-9/{}_.-]+)['""]?:$") {
            $CurrentPath = $matches[1]
            $CurrentMethod = "Unknown"
            continue
        }

        # B) METHOD DETECTION: get:, post:, etc.
        if ($Trimmed -match "^(get|post|put|patch|delete):$") {
            $CurrentMethod = $matches[1].ToUpper()
            continue
        }

        # C) PERMISSION DETECTION: "- Word.Word" pattern
        if ($Trimmed -match "^-\s+([A-Z][a-zA-Z0-9]+(\.[a-zA-Z0-9]+)+)$") {
            $PossiblePerm = $matches[1]

            if ($CurrentPath -ne "Unknown" -and $CurrentMethod -ne "Unknown") {
                if ($PossiblePerm -notmatch "microsoft.graph" -and $PossiblePerm.Length -lt 60) {
                    $AllResults.Add([PSCustomObject]@{
                        Permission = $PossiblePerm
                        Method     = $CurrentMethod
                        Endpoint   = $CurrentPath
                        ApiVersion = $Api.Name
                    })
                }
            }
        }
    }

    $Reader.Close()
    Write-Host "   -> $Counter total lines scanned." -ForegroundColor DarkGray
}

Write-Host ""

# 3. Save Results
if ($AllResults.Count -gt 0) {
    # Remove duplicates
    $UniqueResults = $AllResults | Sort-Object Permission, Method, Endpoint, ApiVersion -Unique
    
    $UniqueResults | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
    
    # Statistics
    $V1Count = ($UniqueResults | Where-Object { $_.ApiVersion -eq "v1.0" }).Count
    $BetaCount = ($UniqueResults | Where-Object { $_.ApiVersion -eq "beta" }).Count
    $UniquePerms = ($UniqueResults | Select-Object -ExpandProperty Permission -Unique).Count
    
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "OPERATION COMPLETED!" -ForegroundColor Green
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "File: $OutCsv" -ForegroundColor Yellow
    Write-Host "Total Records: $($UniqueResults.Count)" -ForegroundColor Yellow
    Write-Host "  - v1.0 API: $V1Count" -ForegroundColor White
    Write-Host "  - beta API: $BetaCount" -ForegroundColor White
    Write-Host "Unique Permissions: $UniquePerms" -ForegroundColor Yellow
    Write-Host "================================================" -ForegroundColor Cyan
}
else {
    Write-Error "Files were read but no Permission patterns were found."
}
