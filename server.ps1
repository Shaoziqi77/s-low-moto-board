# 基于 TcpListener 的简易 HTTP/1.1 静态服务器 (S低L3 本地预览)
# 为浏览器兼容性优化：标准响应头 + Connection: close + 明确 Content-Length
param(
  [int]$Port = 8081,
  [string]$Root
)
Add-Type -AssemblyName System.Web
if ([string]::IsNullOrWhiteSpace($Root)) {
  $Root = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$Root = (Resolve-Path -LiteralPath $Root).Path

$mime = @{
  ".html"="text/html; charset=utf-8"
  ".htm"="text/html; charset=utf-8"
  ".js"="application/javascript; charset=utf-8"
  ".mjs"="application/javascript; charset=utf-8"
  ".css"="text/css; charset=utf-8"
  ".json"="application/json; charset=utf-8"
  ".png"="image/png"
  ".jpg"="image/jpeg"
  ".jpeg"="image/jpeg"
  ".gif"="image/gif"
  ".svg"="image/svg+xml"
  ".ico"="image/x-icon"
  ".woff"="font/woff"
  ".woff2"="font/woff2"
  ".ttf"="font/ttf"
  ".txt"="text/plain; charset=utf-8"
  ".md"="text/markdown; charset=utf-8"
  ".csv"="text/csv; charset=utf-8"
}

$utf8 = [System.Text.Encoding]::UTF8
$buf = New-Object byte[] 8192
$CRLF = "`r`n"

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
$listener.Server.ReceiveBufferSize = 65536
$listener.Start()
Write-Output "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Serving $Root on http://127.0.0.1:$Port/"
[Console]::Out.Flush()

while ($true) {
  $client = $null
  $ns = $null
  $fs = $null
  try {
    $client = $listener.AcceptTcpClient()
    $client.ReceiveTimeout = 5000
    $client.SendTimeout = 5000
    $ns = $client.GetStream()

    # 读取请求头
    $rms = New-Object System.IO.MemoryStream
    $headerDone = $false
    $headerLen = 0
    $stopAt = (Get-Date).AddSeconds(3)
    while (-not $headerDone -and (Get-Date) -lt $stopAt) {
      try {
        $read = $ns.Read($buf, 0, $buf.Length)
      } catch { $read = 0 }
      if ($read -le 0) { break }
      $rms.Write($buf, 0, $read)
      $headerLen += $read
      $raw = $utf8.GetString($rms.ToArray())
      if ($raw.Contains($CRLF + $CRLF) -or $raw.Contains("`n`n")) { $headerDone = $true }
      if ($headerLen -gt 32768) { break }
    }
    $requestText = $utf8.GetString($rms.ToArray())
    $rms.Dispose()

    # 解析请求行
    $path = "/"
    $firstLine = ($requestText -split "`r`n")[0]
    if ($firstLine -match '^[A-Z]+\s+(\S+)\s+HTTP/') {
      $path = $Matches[1]
    }
    $path = [System.Web.HttpUtility]::UrlDecode($path)
    $qIdx = $path.IndexOf('?')
    if ($qIdx -ge 0) { $path = $path.Substring(0, $qIdx) }
    if ($path -eq "/" -or $path -eq "" -or $path -eq "/favicon.ico" -and -not (Test-Path -LiteralPath (Join-Path $Root "favicon.ico") -PathType Leaf)) {
      if ($path -eq "/" -or $path -eq "") { $path = "/index.html" }
    }
    $safe = $path.TrimStart('/','\').Replace('/',[char]92)
    $filePath = Join-Path $Root $safe

    $statusCode = 200
    $statusText = "OK"
    $contentType = "application/octet-stream"
    $bodyPath = $null

    if (Test-Path -LiteralPath $filePath -PathType Leaf) {
      $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
      if ($mime.ContainsKey($ext)) { $contentType = $mime[$ext] }
      $bodyPath = $filePath
    } else {
      $statusCode = 404
      $statusText = "Not Found"
      $contentType = "text/html; charset=utf-8"
      $body404 = "<!doctype html><html><head><meta charset=`"utf-8`"><title>404</title></head><body style=`"font-family:sans-serif;padding:24px`"><h1>404 Not Found</h1><code>$(($filePath -replace '<','&lt;'))</code></body></html>"
      $bodyPath = Join-Path $Root ("__404_" + [guid]::NewGuid().ToString("N") + ".html")
      [System.IO.File]::WriteAllText($bodyPath, $body404, $utf8)
    }

    $bodyLen = if ($bodyPath -and (Test-Path -LiteralPath $bodyPath -PathType Leaf)) {
      (Get-Item -LiteralPath $bodyPath).Length
    } else { 0 }

    # 发送响应头
    $headerText = "HTTP/1.1 $statusCode $statusText" + $CRLF +
                  "Content-Type: $contentType" + $CRLF +
                  "Content-Length: $bodyLen" + $CRLF +
                  "Cache-Control: no-cache, no-store, must-revalidate" + $CRLF +
                  "Pragma: no-cache" + $CRLF +
                  "Connection: close" + $CRLF +
                  "Date: " + [DateTime]::UtcNow.ToString("r") + $CRLF +
                  "Server: S-LL3-Preview" + $CRLF + $CRLF
    $headerBytes = $utf8.GetBytes($headerText)
    $ns.Write($headerBytes, 0, $headerBytes.Length)
    $ns.Flush()

    # 发送响应体
    if ($bodyLen -gt 0) {
      $fs = [System.IO.File]::OpenRead($bodyPath)
      while ($true) {
        $read = $fs.Read($buf, 0, $buf.Length)
        if ($read -le 0) { break }
        $ns.Write($buf, 0, $read)
      }
      $ns.Flush()
      $fs.Close()
      if ($bodyPath -like "*__404_*") {
        try { [System.IO.File]::Delete($bodyPath) } catch {}
      }
    }

    # 刷新并等待发送完成
    try {
      $ns.Flush()
      Start-Sleep -Milliseconds 20
    } catch {}
  } catch {
    # 静默失败，防止异常打印阻塞主循环
  } finally {
    try { if ($ns) { $ns.Close() } } catch {}
    try { if ($client) { $client.Close() } } catch {}
  }
}
