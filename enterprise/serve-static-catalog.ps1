param(
    [int]$Port = 8090,
    [string]$BindAddress = "127.0.0.1",
    [string]$Root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$resolvedRoot = (Resolve-Path $Root).Path
$ipAddress = [System.Net.IPAddress]::Parse($BindAddress)
$listener = [System.Net.Sockets.TcpListener]::new($ipAddress, $Port)

function Get-ContentType {
    param([string]$Path)

    switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        ".css" { return "text/css; charset=utf-8" }
        ".html" { return "text/html; charset=utf-8" }
        ".js" { return "application/javascript; charset=utf-8" }
        ".json" { return "application/json; charset=utf-8" }
        ".md" { return "text/markdown; charset=utf-8" }
        ".txt" { return "text/plain; charset=utf-8" }
        ".svg" { return "image/svg+xml" }
        ".png" { return "image/png" }
        ".jpg" { return "image/jpeg" }
        ".jpeg" { return "image/jpeg" }
        ".gif" { return "image/gif" }
        ".webp" { return "image/webp" }
        ".woff" { return "font/woff" }
        ".woff2" { return "font/woff2" }
        ".ttf" { return "font/ttf" }
        ".plugin" { return "application/octet-stream" }
        default { return "application/octet-stream" }
    }
}

function Send-Response {
    param(
        [System.Net.Sockets.NetworkStream]$Stream,
        [int]$StatusCode,
        [string]$StatusText,
        [byte[]]$Body = [byte[]]::new(0),
        [string]$ContentType = "text/plain; charset=utf-8",
        [bool]$HeadOnly = $false
    )

    $headers = @(
        "HTTP/1.1 $StatusCode $StatusText",
        "Access-Control-Allow-Origin: *",
        "Access-Control-Allow-Methods: GET, HEAD, OPTIONS",
        "Access-Control-Allow-Headers: *",
        "Cache-Control: no-cache",
        "Content-Type: $ContentType",
        "Content-Length: $($Body.Length)",
        "Connection: close",
        "",
        ""
    ) -join "`r`n"

    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($headers)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    if (-not $HeadOnly -and $Body.Length -gt 0) {
        $Stream.Write($Body, 0, $Body.Length)
    }
}

function Resolve-RequestPath {
    param([string]$RawPath)

    $pathOnly = ($RawPath -split "\?", 2)[0]
    if ($pathOnly -eq "/health") {
        return "__health__"
    }
    $decoded = [System.Uri]::UnescapeDataString($pathOnly).TrimStart("/")
    $decoded = $decoded -replace "/", [System.IO.Path]::DirectorySeparatorChar
    $candidate = [System.IO.Path]::GetFullPath((Join-Path $resolvedRoot $decoded))
    $rootPrefix = $resolvedRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    if ($candidate -ne $resolvedRoot -and -not $candidate.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }
    if ([System.IO.Directory]::Exists($candidate)) {
        $indexPath = Join-Path $candidate "index.html"
        if ([System.IO.File]::Exists($indexPath)) {
            return $indexPath
        }
    }
    return $candidate
}

$listener.Start()
Write-Host "Serving $resolvedRoot at http://$BindAddress`:$Port/ with CORS enabled"
Write-Host "Press Ctrl+C to stop."

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        try {
            $stream = $client.GetStream()
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::ASCII, $false, 1024, $true)
            $requestLine = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($requestLine)) {
                Send-Response -Stream $stream -StatusCode 400 -StatusText "Bad Request"
                continue
            }

            while (-not [string]::IsNullOrEmpty($reader.ReadLine())) {}

            $parts = $requestLine.Split(" ")
            $method = $parts[0].ToUpperInvariant()
            $rawPath = if ($parts.Length -gt 1) { $parts[1] } else { "/" }
            Write-Host ("[{0}] {1} {2}" -f (Get-Date -Format "HH:mm:ss"), $method, $rawPath)

            if ($method -eq "OPTIONS") {
                Send-Response -Stream $stream -StatusCode 204 -StatusText "No Content"
                continue
            }
            if ($method -ne "GET" -and $method -ne "HEAD") {
                $body = [System.Text.Encoding]::UTF8.GetBytes("Method not allowed")
                Send-Response -Stream $stream -StatusCode 405 -StatusText "Method Not Allowed" -Body $body
                continue
            }

            $resolvedPath = Resolve-RequestPath -RawPath $rawPath
            if ($resolvedPath -eq "__health__") {
                $body = [System.Text.Encoding]::UTF8.GetBytes("ok")
                Send-Response -Stream $stream -StatusCode 200 -StatusText "OK" -Body $body -HeadOnly:($method -eq "HEAD")
                continue
            }
            if (-not $resolvedPath -or -not [System.IO.File]::Exists($resolvedPath)) {
                $body = [System.Text.Encoding]::UTF8.GetBytes("Not found")
                Send-Response -Stream $stream -StatusCode 404 -StatusText "Not Found" -Body $body -HeadOnly:($method -eq "HEAD")
                continue
            }

            $bodyBytes = [System.IO.File]::ReadAllBytes($resolvedPath)
            Send-Response `
                -Stream $stream `
                -StatusCode 200 `
                -StatusText "OK" `
                -Body $bodyBytes `
                -ContentType (Get-ContentType -Path $resolvedPath) `
                -HeadOnly:($method -eq "HEAD")
        } catch {
            try {
                $body = [System.Text.Encoding]::UTF8.GetBytes($_.Exception.Message)
                Send-Response -Stream $stream -StatusCode 500 -StatusText "Internal Server Error" -Body $body
            } catch {
            }
        } finally {
            $client.Close()
        }
    }
} finally {
    $listener.Stop()
}
