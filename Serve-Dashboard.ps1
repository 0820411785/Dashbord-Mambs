param(
    [int]$Port = 8080,
    [string]$Root = ".",
    [string]$DefaultFile = "Dashboard_Incidents_12_Juin_2026.html"
)

$ErrorActionPreference = "Stop"
$rootPath = (Resolve-Path -LiteralPath $Root).Path
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, $Port)
$listener.Start()

Write-Output "Dashboard server started on port $Port"
Write-Output "Root: $rootPath"
Write-Output "Press Ctrl+C to stop."

function Get-ContentType {
    param([string]$Path)
    switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        ".html" { "text/html; charset=utf-8" }
        ".htm"  { "text/html; charset=utf-8" }
        ".csv"  { "text/csv; charset=utf-8" }
        ".css"  { "text/css; charset=utf-8" }
        ".js"   { "application/javascript; charset=utf-8" }
        default { "application/octet-stream" }
    }
}

function Send-Response {
    param(
        [System.Net.Sockets.NetworkStream]$Stream,
        [int]$StatusCode,
        [string]$StatusText,
        [byte[]]$Body,
        [string]$ContentType = "text/plain; charset=utf-8"
    )

    $header = "HTTP/1.1 $StatusCode $StatusText`r`nContent-Type: $ContentType`r`nContent-Length: $($Body.Length)`r`nConnection: close`r`n`r`n"
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    $Stream.Write($Body, 0, $Body.Length)
}

while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
        $stream = $client.GetStream()
        $buffer = New-Object byte[] 4096
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { continue }

        $request = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $read)
        $requestLine = ($request -split "`r?`n")[0]
        $parts = $requestLine -split " "
        $urlPath = if ($parts.Count -ge 2) { $parts[1] } else { "/" }
        $urlPath = [System.Uri]::UnescapeDataString(($urlPath -split "\?")[0])

        $fileName = if ($urlPath -eq "/" -or [string]::IsNullOrWhiteSpace($urlPath)) {
            $DefaultFile
        } else {
            [System.IO.Path]::GetFileName($urlPath)
        }

        $filePath = Join-Path $rootPath $fileName
        $fullPath = [System.IO.Path]::GetFullPath($filePath)

        if (-not $fullPath.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            $body = [System.Text.Encoding]::UTF8.GetBytes("Fichier introuvable.")
            Send-Response $stream 404 "Not Found" $body
            continue
        }

        $bodyBytes = [System.IO.File]::ReadAllBytes($fullPath)
        Send-Response $stream 200 "OK" $bodyBytes (Get-ContentType $fullPath)
    }
    catch {
        try {
            $body = [System.Text.Encoding]::UTF8.GetBytes("Erreur serveur: $($_.Exception.Message)")
            Send-Response $stream 500 "Internal Server Error" $body
        } catch {}
    }
    finally {
        $client.Close()
    }
}
