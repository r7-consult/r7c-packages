param(
    [string]$SourceRoot = "",
    [string]$DestinationRoot = "",
    [int]$TimeoutSeconds = 1800
)

$ErrorActionPreference = "Stop"

if (-not $SourceRoot) {
    $SourceRoot = Split-Path -Parent $PSScriptRoot
}

if (-not $DestinationRoot) {
    $programFiles = if ($env:ProgramFiles) { $env:ProgramFiles } else { "C:\Program Files" }
    $DestinationRoot = Join-Path $programFiles "R7-Office\Editors\editors\sdkjs-plugins\{8455F88E-A130-4EE3-8465-FC2B2A810360}"
}

$files = @(
    "index.html",
    "styles\style.css",
    "scripts\main.js",
    "scripts\ui\r7chat_ui.js",
    "scripts\platform\openrouter_client.js",
    "scripts\platform\r7chat_http_client.js",
    "scripts\features\settings\r7chat_settings_service.js",
    "scripts\features\settings\r7chat_settings_panel.js",
    "scripts\features\chat\r7chat_chat_runtime.js"
)

function Test-FileUnlocked {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $true
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        return $true
    } catch {
        return $false
    } finally {
        if ($stream) {
            $stream.Dispose()
        }
    }
}

function Sync-Files {
    foreach ($relativePath in $files) {
        $sourcePath = Join-Path $SourceRoot $relativePath
        $destinationPath = Join-Path $DestinationRoot $relativePath
        $destinationDir = Split-Path -Path $destinationPath -Parent

        if (-not (Test-Path -LiteralPath $sourcePath)) {
            throw "Source file not found: $sourcePath"
        }

        if (-not (Test-Path -LiteralPath $destinationDir)) {
            New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
        }

        Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
    }
}

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)

while ((Get-Date) -lt $deadline) {
    $locked = $false
    foreach ($relativePath in $files) {
        $destinationPath = Join-Path $DestinationRoot $relativePath
        if (-not (Test-FileUnlocked -Path $destinationPath)) {
            $locked = $true
            break
        }
    }

    if (-not $locked) {
        Sync-Files
        Write-Output "Installed plugin synced successfully."
        exit 0
    }

    Start-Sleep -Seconds 2
}

throw "Timed out waiting for installed plugin files to unlock."
