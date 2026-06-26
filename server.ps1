# 轻量本地 HTTP 静态服务器 —— 用于 S低L3-3 数据处理工具预览
param(
  [int]$Port = 8080,
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
  ".xlsx"="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
}

$utf8 = [System.Text.Encoding]::UTF8
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
try {
  $listener.Start()
} catch {
  Write-Output "Failed to start server on port $Port`: $_"
  exit 1
}
Write-Output "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Serving $Root on http://127.0.0.1:$Port/"
Write-Output "Press Ctrl+C to stop."

try {
  while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $res = $ctx.Response
    try {
      $path = $req.Url.LocalPath
      $path = [System.Web.HttpUtility]::UrlDecode($path)
      if ($path -eq "/" -or $path -eq "") { $path = "/index.html" }
      $safe = $path.TrimStart('/','\').Replace('/',[char]92)
      $filePath = Join-Path $Root $safe

      $statusCode = 200
      $statusText = "OK"
      $contentType = "application/octet-stream"
      $bodyBytes = $null

      if (Test-Path -LiteralPath $filePath -PathType Leaf) {
        $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
        if ($mime.ContainsKey($ext)) { $contentType = $mime[$ext] }
        try {
          $bodyBytes = [System.IO.File]::ReadAllBytes($filePath)
        } catch {
          $statusCode = 500; $statusText = "Server Error"
        }
      } else {
        $statusCode = 404
        $statusText = "Not Found"
        $contentType = "text/html; charset=utf-8"
        $bodyBytes = $utf8.GetBytes("<!doctype html><html><head><meta charset='utf-8'><title>404</title></head><body style='font-family:sans-serif;padding:24px'><h1>404 Not Found</h1><code>$([System.Net.WebUtility]::HtmlEncode($filePath))</code></body></html>")
      }

      $res.StatusCode = $statusCode
      $res.StatusDescription = $statusText
      $res.ContentType = $contentType
      if ($bodyBytes -ne $null) { $res.ContentLength64 = $bodyBytes.Length }
      $res.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate")
      $res.AddHeader("Pragma", "no-cache")
      try {
        if ($bodyBytes -ne $null) { $res.OutputStream.Write($bodyBytes, 0, $bodyBytes.Length) }
        $res.OutputStream.Flush()
      } catch {}
    } finally {
      try { $res.Close() } catch {}
    }
  }
} finally {
  try { $listener.Stop() } catch {}
  try { $listener.Close() } catch {}
  Write-Output "Server stopped."
}
