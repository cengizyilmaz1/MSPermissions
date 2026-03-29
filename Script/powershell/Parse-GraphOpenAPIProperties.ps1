<#
    .SYNOPSIS
    Optimized Graph Resource Schema Extractor (High Performance)
    
    .DESCRIPTION
    Extracts Microsoft Graph API resource schemas from OpenAPI specifications.
    Uses Hash Maps for O(1) lookups and Memoization to prevent re-parsing inherited schemas.
    Supports parallel processing for v1.0 and beta versions.
    
    .PARAMETER OutputPath
    The directory to save the output JSON file. Defaults to ".\data"
    
    .PARAMETER Force
    Force re-download of OpenAPI YAML files even if cached
    
    .EXAMPLE
    .\Script\powershell\Parse-GraphOpenAPIProperties.ps1 -OutputPath ".\data" -Force
#>

param (
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'

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

$RepoRoot = if ($PSScriptRoot) {
    Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
} else {
    (Get-Location).Path
}

if (-not $OutputPath) {
    $OutputPath = Join-Path $RepoRoot "data"
}

# Performance tracking
$script:StartTime = Get-Date

Write-Host "`n[Graph OpenAPI Schema Extractor v2.0]`n" -ForegroundColor Cyan

# --- 1. MODULE CHECK ---
function Ensure-YamlModule {
    if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
        Write-Host "Installing powershell-yaml module..." -ForegroundColor Yellow
        try {
            Install-Module -Name powershell-yaml -Force -Scope CurrentUser -AllowClobber
            Write-Host "   Module installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install powershell-yaml module: $_"
            exit 1
        }
    }
    Import-Module powershell-yaml -ErrorAction Stop
}

# --- 2. SCHEMA CACHE (Thread-safe for potential parallel processing) ---
$script:SchemaCache = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
$script:EntityDetectionCache = [System.Collections.Concurrent.ConcurrentDictionary[string, bool]]::new()
$script:MethodIndex = @{}

# --- 3. HELPER FUNCTIONS ---

function Parse-OpenAPISchema {
    param(
        [hashtable]$Schema,
        [string]$SchemaName,
        [hashtable]$AllSchemas,
        [string]$Version
    )

    # Cache Key: Version + SchemaName
    $cacheKey = "$Version::$SchemaName"
    
    $cachedResult = $null
    if ($script:SchemaCache.TryGetValue($cacheKey, [ref]$cachedResult)) {
        return $cachedResult
    }

    $properties = [System.Collections.ArrayList]::new()

    # 1. Handle inheritance (allOf) - Recursive
    if ($Schema.allOf) {
        foreach ($item in $Schema.allOf) {
            if ($item.'$ref') {
                $refName = $item.'$ref' -replace '#/components/schemas/', ''
                if ($AllSchemas.ContainsKey($refName)) {
                    $inheritedProps = Parse-OpenAPISchema -Schema $AllSchemas[$refName] -SchemaName $refName -AllSchemas $AllSchemas -Version $Version
                    if ($inheritedProps) {
                        foreach ($prop in @($inheritedProps)) {
                            [void]$properties.Add($prop)
                        }
                    }
                }
            }
            # Inline allOf properties
            if ($item.properties) {
                $extracted = Extract-PropsFromHashtable -PropHash $item.properties
                if ($extracted) {
                    foreach ($prop in @($extracted)) {
                        [void]$properties.Add($prop)
                    }
                }
            }
        }
    }

    # 2. Direct properties
    if ($Schema.properties) {
        $directProps = Extract-PropsFromHashtable -PropHash $Schema.properties
        if ($directProps) {
            foreach ($prop in @($directProps)) {
                [void]$properties.Add($prop)
            }
        }
    }

    # Remove duplicates (keep last occurrence for override behavior)
    $uniqueProps = $properties | Group-Object -Property name | ForEach-Object { $_.Group[-1] }
    
    # Cache the result
    [void]$script:SchemaCache.TryAdd($cacheKey, $uniqueProps)
    
    return $uniqueProps
}

function Test-IsEntitySchema {
    param(
        [string]$SchemaName,
        [hashtable]$AllSchemas,
        [System.Collections.Generic.HashSet[string]]$Visited = $null
    )

    if (-not $SchemaName) {
        return $false
    }

    $cached = $false
    if ($script:EntityDetectionCache.TryGetValue($SchemaName, [ref]$cached)) {
        return $cached
    }

    if (-not $Visited) {
        $Visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    if (-not $Visited.Add($SchemaName)) {
        return $false
    }

    $result = $false
    $cleanName = $SchemaName -replace '^microsoft\.graph\.', ''

    if ($SchemaName -match '(^|\\.)entity$') {
        $result = $true
    }
    elseif ($AllSchemas.ContainsKey($SchemaName)) {
        $schema = $AllSchemas[$SchemaName]

        if ($schema.properties -and $schema.properties.ContainsKey('id')) {
            $result = $true
        }
        elseif ($schema.allOf) {
            foreach ($item in $schema.allOf) {
                if ($item.'$ref') {
                    $refName = $item.'$ref' -replace '#/components/schemas/', ''
                    if (Test-IsEntitySchema -SchemaName $refName -AllSchemas $AllSchemas -Visited $Visited) {
                        $result = $true
                        break
                    }
                }
            }
        }

        if (-not $result -and $script:MethodIndex.ContainsKey($cleanName.ToLower())) {
            $result = $true
        }
    }

    [void]$script:EntityDetectionCache.TryAdd($SchemaName, $result)
    return $result
}

function Extract-PropsFromHashtable {
    param([hashtable]$PropHash)
    
    if (-not $PropHash -or $PropHash.Count -eq 0) {
        return $null
    }
    
    $results = [System.Collections.ArrayList]::new()
    
    foreach ($key in $PropHash.Keys) {
        $p = $PropHash[$key]
        $type = "object"
        
        # Type determination (optimized)
        if ($p.'$ref') { 
            $type = ($p.'$ref' -split '/')[-1]
        }
        elseif ($p.type -eq 'array') {
            if ($p.items.'$ref') {
                $subType = ($p.items.'$ref' -split '/')[-1]
                $type = "$subType collection"
            } 
            elseif ($p.items.type) {
                $type = "$($p.items.type) collection"
            }
            else {
                $type = "array"
            }
        }
        elseif ($p.format) { 
            $type = $p.format 
        }
        elseif ($p.type) { 
            $type = $p.type 
        }
        elseif ($p.enum) {
            $type = "enum"
        }

        [void]$results.Add([PSCustomObject]@{
            name        = $key
            type        = $type
            description = if ($p.description) { $p.description.Trim() } else { "" }
            readOnly    = [bool]$p.readOnly
            nullable    = [bool]$p.nullable
        })
    }
    
    return $results
}

function Build-MethodIndex {
    param([hashtable]$Paths)
    
    if (-not $Paths -or $Paths.Count -eq 0) {
        return @{}
    }
    
    $index = @{}
    $httpVerbs = @('get', 'post', 'patch', 'put', 'delete')

    foreach ($pathUrl in $Paths.Keys) {
        # Extract primary resource from URL path
        if ($pathUrl -match '^/([^/{]+)') {
            $segment = $matches[1]
            
            # Singularize resource name (simple approach)
            $resourceName = if ($segment -match "ies$") {
                $segment -replace "ies$", "y"
            }
            elseif ($segment -match "ses$") {
                $segment -replace "ses$", "s"
            }
            elseif ($segment -match "s$" -and $segment -notmatch "(ss|us|is)$") {
                $segment.Substring(0, $segment.Length - 1)
            }
            else {
                $segment
            }

            if (-not $index.ContainsKey($resourceName)) {
                $index[$resourceName] = [System.Collections.ArrayList]::new()
            }

            $pathItem = $Paths[$pathUrl]
            foreach ($verb in $httpVerbs) {
                if ($pathItem[$verb]) {
                    $op = $pathItem[$verb]
                    $simpleMethod = switch ($verb) {
                        'get'    { if ($pathUrl -match '{[^}]+}$') { "Get" } else { "List" } }
                        'post'   { "Create" }
                        'patch'  { "Update" }
                        'put'    { "Replace" }
                        'delete' { "Delete" }
                    }

                    [void]$index[$resourceName].Add([PSCustomObject]@{
                        method      = $simpleMethod
                        httpMethod  = $verb.ToUpper()
                        path        = $pathUrl
                        description = if ($op.summary) { $op.summary } elseif ($op.description) { $op.description.Split("`n")[0] } else { "" }
                        operationId = $op.operationId
                    })
                }
            }
        }
    }
    
    return $index
}

function Get-OpenAPIYaml {
    param(
        [string]$Version,
        [string]$Url,
        [switch]$Force
    )
    
    $cacheDir = Get-OpenApiCacheDirectory
    
    $tempFile = Join-Path $cacheDir "openapi-$Version.yaml"
    $metaFile = Join-Path $cacheDir "openapi-$Version.meta"
    
    # Check cache validity (24 hours)
    $cacheValid = $false
    if ((Test-Path $tempFile) -and (Test-Path $metaFile) -and -not $Force) {
        $meta = Get-Content $metaFile | ConvertFrom-Json
        $cacheAge = (Get-Date) - [DateTime]$meta.downloaded
        if ($cacheAge.TotalHours -lt 24) {
            $cacheValid = $true
            Write-Host "   Using cached YAML (age: $([math]::Round($cacheAge.TotalHours, 1))h)" -ForegroundColor Gray
        }
    }
    
    if (-not $cacheValid) {
        Write-Host "   Downloading OpenAPI YAML..." -ForegroundColor Yellow
        try {
            Invoke-WebRequest -Uri $Url -OutFile $tempFile -UseBasicParsing
            @{ downloaded = (Get-Date).ToString("o"); url = $Url } | ConvertTo-Json | Set-Content $metaFile
            Write-Host "   Downloaded successfully" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to download $Version OpenAPI spec: $_"
            if (Test-Path $tempFile) {
                Write-Host "   Using stale cache" -ForegroundColor Yellow
            }
            else {
                return $null
            }
        }
    }
    
    Write-Host "   Parsing YAML (this may take a moment)..." -ForegroundColor Yellow
    try {
        $content = Get-Content $tempFile -Raw
        $parsed = ConvertFrom-Yaml $content
        Write-Host "   Parsed successfully" -ForegroundColor Green
        return $parsed
    }
    catch {
        Write-Error "YAML parsing failed for $Version : $_"
        return $null
    }
}

function Process-OpenAPIVersion {
    param(
        [string]$Version,
        [hashtable]$YamlData,
        [hashtable]$FinalOutput
    )
    
    # Clear version-specific cache entries
    $keysToRemove = $script:SchemaCache.Keys | Where-Object { $_ -like "$Version::*" }
    foreach ($key in $keysToRemove) {
        [void]$script:SchemaCache.TryRemove($key, [ref]$null)
    }
    
    $Schemas = $YamlData.components.schemas
    if (-not $Schemas) {
        Write-Warning "No schemas found in $Version"
        return $FinalOutput
    }
    
    # Build method index for O(1) lookups
    Write-Host "   Building method index..." -ForegroundColor Gray
    $MethodIndex = Build-MethodIndex -Paths $YamlData.paths
    $script:MethodIndex = $MethodIndex
    Write-Host "   Indexed $($MethodIndex.Count) resources" -ForegroundColor Gray
    
    Write-Host "   Extracting entity schemas..." -ForegroundColor Gray
    
    $processedCount = 0
    $skippedCount = 0
    $totalSchemas = $Schemas.Keys.Count
    
    foreach ($schemaName in $Schemas.Keys) {
        # Skip internal/helper types
        if ($schemaName -match '^(odata\.|microsoft\.graph\.odata|CollectionResponse|ODataErrors|StringCollection|Int32Collection)') { 
            $skippedCount++
            continue 
        }
        
        $schema = $Schemas[$schemaName]
        
        # Entity detection: recursive inheritance or method-indexed top-level resource
        $isEntity = Test-IsEntitySchema -SchemaName $schemaName -AllSchemas $Schemas

        if (-not $isEntity) { 
            $skippedCount++
            continue 
        }

        # Clean name (remove namespace)
        $cleanName = $schemaName -replace '^microsoft\.graph\.', ''

        # Initialize or update output object
        if (-not $FinalOutput.ContainsKey($cleanName)) {
            $FinalOutput[$cleanName] = @{
                name        = $cleanName
                description = if ($schema.description) { $schema.description.Trim() } else { "" }
                properties  = @{ 'v1.0' = @(); 'beta' = @() }
                methods     = @{ 'v1.0' = @(); 'beta' = @() }
            }
        }
        elseif ($schema.description -and -not $FinalOutput[$cleanName].description) {
            $FinalOutput[$cleanName].description = $schema.description.Trim()
        }

        # Extract properties (uses memoization)
        $props = Parse-OpenAPISchema -Schema $schema -SchemaName $schemaName -AllSchemas $Schemas -Version $Version
        $FinalOutput[$cleanName].properties[$Version] = $props

        # Get methods from index (O(1) lookup)
        $lowerName = $cleanName.ToLower()
        if ($MethodIndex.ContainsKey($lowerName)) {
            $FinalOutput[$cleanName].methods[$Version] = $MethodIndex[$lowerName]
        }
        
        $processedCount++
    }

    Write-Host "   Processed $processedCount entities (skipped $skippedCount non-entities)" -ForegroundColor Green
    
    return $FinalOutput
}

# --- MAIN EXECUTION ---

try {
    # Ensure required module
    Ensure-YamlModule
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
    }
    
    $openApiUrls = @{
        'v1.0' = 'https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/v1.0/openapi.yaml'
        'beta' = 'https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/beta/openapi.yaml'
    }
    
    $FinalOutput = @{}
    
    foreach ($version in @('v1.0', 'beta')) {
        Write-Host "Processing [$version]..." -ForegroundColor Cyan
        
        $yamlData = Get-OpenAPIYaml -Version $version -Url $openApiUrls[$version] -Force:$Force
        
        if ($yamlData) {
            $FinalOutput = Process-OpenAPIVersion -Version $version -YamlData $yamlData -FinalOutput $FinalOutput
        }
        
        # Memory cleanup
        $yamlData = $null
        [System.GC]::Collect()
    }
    
    # Filter out entities with no properties
    Write-Host "`nFinalizing output..." -ForegroundColor Cyan
    $CleanOutput = @{}
    foreach ($key in $FinalOutput.Keys) {
        $item = $FinalOutput[$key]
        $hasV1Props = $item.properties['v1.0'] -and $item.properties['v1.0'].Count -gt 0
        $hasBetaProps = $item.properties['beta'] -and $item.properties['beta'].Count -gt 0
        if ($hasV1Props -or $hasBetaProps) {
            $CleanOutput[$key] = $item
        }
    }
    
    # Save output
    $outputFile = Join-Path $OutputPath "GraphResourceSchemas.json"
    Write-Host "   Saving to $outputFile..." -ForegroundColor Gray
    
    $CleanOutput | ConvertTo-Json -Depth 6 -Compress | Out-File $outputFile -Encoding UTF8
    
    # Statistics
    $v1Count = ($CleanOutput.Values | Where-Object { $_.properties['v1.0'].Count -gt 0 }).Count
    $betaCount = ($CleanOutput.Values | Where-Object { $_.properties['beta'].Count -gt 0 }).Count
    $duration = (Get-Date) - $script:StartTime
    
    Write-Host "`nExtraction complete" -ForegroundColor Green
    Write-Host "Total Resources: $($CleanOutput.Count)" -ForegroundColor Green
    Write-Host "v1.0 Resources:  $v1Count" -ForegroundColor Green
    Write-Host "Beta Resources:  $betaCount" -ForegroundColor Green
    Write-Host "Duration:        $([math]::Round($duration.TotalSeconds, 1)) seconds" -ForegroundColor Green
    Write-Host "Output:          $outputFile" -ForegroundColor Green
}
catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    exit 1
}
finally {
    # Cleanup
    $script:SchemaCache.Clear()
    $script:EntityDetectionCache.Clear()
    [System.GC]::Collect()
}
