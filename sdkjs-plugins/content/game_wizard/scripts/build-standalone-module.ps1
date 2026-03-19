param(
    [Parameter(Mandatory = $true)]
    [string]$ModuleName
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$moduleDir = Join-Path $repoRoot (Join-Path 'modules' $ModuleName)

if (-not (Test-Path -LiteralPath $moduleDir)) {
    throw "Module folder not found: $moduleDir"
}

$configPath = Join-Path $moduleDir 'config.json'
if (-not (Test-Path -LiteralPath $configPath)) {
    throw "Module config not found: $configPath"
}

$config = Get-Content -Raw -LiteralPath $configPath | ConvertFrom-Json
$tempRoot = Join-Path $env:TEMP ($ModuleName + '_' + [Guid]::NewGuid().ToString())
$outFile = Join-Path $moduleDir ($ModuleName + '.plugin')

function Get-RelativeUnixPath([string]$BasePath, [string]$TargetPath) {
    $base = [System.IO.Path]::GetFullPath($BasePath)
    if (-not $base.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $base += [System.IO.Path]::DirectorySeparatorChar
    }
    $target = [System.IO.Path]::GetFullPath($TargetPath)
    $baseUri = [Uri]$base
    $targetUri = [Uri]$target
    return [Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())
}

function Resolve-RepoRelative([string]$RelativePath) {
    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return $RelativePath
    }
    $full = [System.IO.Path]::GetFullPath((Join-Path $moduleDir $RelativePath))
    return (Get-RelativeUnixPath $repoRoot $full).Replace('\', '/')
}

New-Item -ItemType Directory -Path $tempRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'modules') | Out-Null
Copy-Item -LiteralPath $moduleDir -Destination (Join-Path $tempRoot 'modules') -Recurse -Force

if (Test-Path -LiteralPath (Join-Path $repoRoot 'resources')) {
    Copy-Item -LiteralPath (Join-Path $repoRoot 'resources') -Destination $tempRoot -Recurse -Force
}

if (Test-Path -LiteralPath (Join-Path $repoRoot 'styles\game-shell.css')) {
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'styles') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $repoRoot 'styles\game-shell.css') -Destination (Join-Path $tempRoot 'styles') -Force
}

if (Test-Path -LiteralPath (Join-Path $repoRoot 'vendor\oo')) {
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'vendor') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $repoRoot 'vendor\oo') -Destination (Join-Path $tempRoot 'vendor') -Recurse -Force
}

$standaloneConfig = $config | ConvertTo-Json -Depth 20 | ConvertFrom-Json
if ($standaloneConfig.variations) {
    foreach ($variation in $standaloneConfig.variations) {
        if ($variation.url) {
            $variation.url = Resolve-RepoRelative $variation.url
        }
        if ($variation.icons) {
            $resolvedIcons = @()
            foreach ($icon in $variation.icons) {
                if ($icon) {
                    $resolvedIcons += Resolve-RepoRelative $icon
                }
            }
            $variation.icons = $resolvedIcons
        }
    }
}

$standaloneConfig | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $tempRoot 'config.json') -Encoding UTF8

if (Test-Path -LiteralPath $outFile) {
    Remove-Item -LiteralPath $outFile -Force
}

$zipPath = Join-Path $env:TEMP ($ModuleName + '_' + [Guid]::NewGuid().ToString() + '.zip')

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)

try {
    $files = Get-ChildItem -LiteralPath $tempRoot -File -Recurse -Force
    foreach ($file in $files) {
        if ($file.Extension -ieq '.plugin') {
            continue
        }
        $relative = (Get-RelativeUnixPath $tempRoot $file.FullName).Replace('\', '/')
        $parts = $relative -split '/+'
        $skip = $false
        foreach ($part in $parts) {
            if ($part.StartsWith('.')) {
                $skip = $true
                break
            }
        }
        if ($skip) {
            continue
        }
        [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $relative, [IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
}
finally {
    if ($zip) {
        $zip.Dispose()
    }
}

Move-Item -LiteralPath $zipPath -Destination $outFile -Force
Remove-Item -LiteralPath $tempRoot -Recurse -Force
Write-Host ('Standalone module plugin created: ' + $outFile)
