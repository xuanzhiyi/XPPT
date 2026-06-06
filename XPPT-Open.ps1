# XPPT-Open.ps1
# Opens a .pptx file directly in the XPPT Chrome extension viewer.
# Called by Windows file association: XPPT-Open.ps1 "C:\path\to\file.pptx"

param([string]$FilePath)

if (-not $FilePath) {
    Write-Host "Usage: XPPT-Open.ps1 <path-to-file.pptx>"
    exit 1
}

# ── Find the XPPT extension ID from Chrome's installed extensions ────────────
$profiles = @(
    "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Extensions",
    "$env:LOCALAPPDATA\Google\Chrome\User Data\Profile 1\Extensions"
)

$extId = $null
foreach ($extDir in $profiles) {
    if (-not (Test-Path $extDir)) { continue }
    foreach ($idFolder in Get-ChildItem $extDir -Directory -ErrorAction SilentlyContinue) {
        $manifests = Get-ChildItem $idFolder.FullName -Recurse -Filter "manifest.json" -ErrorAction SilentlyContinue
        foreach ($mf in $manifests) {
            $txt = Get-Content $mf.FullName -Raw -ErrorAction SilentlyContinue
            if ($txt -match '"XPPT') {
                $extId = $idFolder.Name
                break
            }
        }
        if ($extId) { break }
    }
    if ($extId) { break }
}

if (-not $extId) {
    [System.Windows.Forms.MessageBox]::Show(
        "XPPT extension not found in Chrome.`n`nPlease install it first:`n1. Open chrome://extensions`n2. Enable Developer mode`n3. Load unpacked → select the XPPT folder",
        "XPPT – Extension not found",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    Start-Process "chrome" "chrome://extensions"
    exit 1
}

# ── Build the viewer URL ─────────────────────────────────────────────────────
$normalized = $FilePath.Replace("\", "/").TrimStart("/")
$fileUrl    = "file:///" + $normalized
$encoded    = [System.Uri]::EscapeDataString($fileUrl)
$name       = [System.Uri]::EscapeDataString([System.IO.Path]::GetFileName($FilePath))
$viewerUrl  = "chrome-extension://$extId/pptx_viewer.html?url=$encoded&name=$name"

# ── Open Chrome ──────────────────────────────────────────────────────────────
$chrome = @(
    "$env:PROGRAMFILES\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
    "${env:PROGRAMFILES(x86)}\Google\Chrome\Application\chrome.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $chrome) {
    Write-Host "Chrome not found."
    exit 1
}

Start-Process $chrome "--new-window `"$viewerUrl`""
