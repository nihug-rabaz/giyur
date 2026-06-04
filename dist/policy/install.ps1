#Requires -RunAsAdministrator
$ErrorActionPreference = "Stop"

$here       = Split-Path -Parent $MyInvocation.MyCommand.Path
$installDir = "C:\ProgramData\Zimunim"
$extId      = "dhlleojgaiclmkkipcooihgffajlochd"
$version    = "1.0.2"

if (-not (Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir | Out-Null }
Copy-Item -Path (Join-Path $here "..\zimunim.crx") -Destination $installDir -Force
Copy-Item -Path (Join-Path $here "update.xml")     -Destination $installDir -Force

$browsers = @("Google\Chrome","BraveSoftware\Brave","Microsoft\Edge")
foreach ($key in $browsers) {
    $base = "HKLM:\SOFTWARE\Policies\$key"
    foreach ($leaf in @("ExtensionInstallForcelist","ExtensionInstallAllowlist","ExtensionInstallSources")) {
        if (-not (Test-Path "$base\$leaf")) { New-Item -Path "$base\$leaf" -Force | Out-Null }
    }
    Set-ItemProperty -Path "$base\ExtensionInstallForcelist" -Name "1" -Value "$extId;file:///$($installDir -replace '\\','/')/update.xml"
    Set-ItemProperty -Path "$base\ExtensionInstallAllowlist" -Name "1" -Value $extId
    Set-ItemProperty -Path "$base\ExtensionInstallSources"   -Name "1" -Value "file:///*"
    Write-Host "Configured policy for $key" -ForegroundColor Green
}

Write-Host ""
Write-Host "Extension $extId v$version installed via policy." -ForegroundColor Green
Write-Host "Restart Chrome/Brave/Edge to load the extension." -ForegroundColor Yellow
