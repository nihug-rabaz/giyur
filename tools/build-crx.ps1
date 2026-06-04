$ErrorActionPreference = "Stop"

$repoRoot  = (Resolve-Path "$PSScriptRoot\..").Path
$distDir   = Join-Path $repoRoot "dist"
$stageDir  = Join-Path $distDir  "extension"
$keyFile   = Join-Path $distDir  "zimunim.pem"
$crxFile   = Join-Path $distDir  "zimunim.crx"
$zipFile   = Join-Path $distDir  "zimunim-unpacked.zip"

$include = @(
    "manifest.json",
    "popup.html",
    "popup.js",
    "content.js",
    "export-page.js",
    "index.umd.js",
    "icons"
)

if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
if (Test-Path $stageDir)       { Remove-Item $stageDir -Recurse -Force }
New-Item -ItemType Directory -Path $stageDir | Out-Null

Write-Host "Staging files..." -ForegroundColor Cyan
foreach ($item in $include) {
    $src = Join-Path $repoRoot $item
    if (Test-Path $src) {
        Copy-Item $src -Destination $stageDir -Recurse -Force
        Write-Host "  + $item"
    } else {
        Write-Warning "  - missing: $item"
    }
}

if (Test-Path $zipFile) { Remove-Item $zipFile -Force }
Compress-Archive -Path "$stageDir\*" -DestinationPath $zipFile -Force
Write-Host "Created ZIP: $zipFile" -ForegroundColor Green

$browserCandidates = @(
    "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "${env:LocalAppData}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles}\BraveSoftware\Brave-Browser\Application\brave.exe",
    "${env:ProgramFiles(x86)}\BraveSoftware\Brave-Browser\Application\brave.exe",
    "${env:LocalAppData}\BraveSoftware\Brave-Browser\Application\brave.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
)
$browser = $browserCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if (-not $browser) {
    Write-Warning "No Chrome/Brave/Edge found. Only the ZIP was created."
    exit 0
}

Write-Host "Packing CRX with: $browser" -ForegroundColor Cyan

if (Test-Path $crxFile)             { Remove-Item $crxFile -Force }
if (Test-Path "$stageDir.crx")      { Remove-Item "$stageDir.crx" -Force }
if (Test-Path "$stageDir.pem")      { Remove-Item "$stageDir.pem" -Force }

$packArgs = @("--pack-extension=$stageDir")
if (Test-Path $keyFile) {
    Write-Host "Re-using existing private key (keeps extension ID stable)" -ForegroundColor Cyan
    $packArgs += "--pack-extension-key=$keyFile"
}

$proc = Start-Process -FilePath $browser -ArgumentList $packArgs -PassThru -WindowStyle Hidden

$generatedCrx = "$stageDir.crx"
$generatedKey = "$stageDir.pem"
$timeoutSec   = 30
$elapsed      = 0
while (-not (Test-Path $generatedCrx) -and $elapsed -lt $timeoutSec) {
    Start-Sleep -Milliseconds 500
    $elapsed += 0.5
}

if (-not (Test-Path $generatedCrx)) {
    Write-Warning "CRX not produced in $timeoutSec sec. The browser may be running and silently completed - check $distDir manually."
} else {
    Move-Item $generatedCrx $crxFile -Force
    Write-Host "Created CRX: $crxFile" -ForegroundColor Green
}

if (Test-Path $generatedKey) {
    Move-Item $generatedKey $keyFile -Force
    Write-Host "Saved private key: $keyFile" -ForegroundColor Yellow
    Write-Host "  -> KEEP THIS FILE. Re-use it on next builds so the extension ID stays the same." -ForegroundColor Yellow
}

if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }

Write-Host ""
Write-Host "==== Output ====" -ForegroundColor Green
Get-ChildItem $distDir -File | Select-Object Name, @{n="Size";e={"{0:N0} B" -f $_.Length}}, LastWriteTime | Format-Table -AutoSize
