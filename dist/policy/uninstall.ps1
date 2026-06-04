#Requires -RunAsAdministrator
$ErrorActionPreference = "Continue"

$browsers = @("Google\Chrome","BraveSoftware\Brave","Microsoft\Edge")
foreach ($key in $browsers) {
    foreach ($leaf in @("ExtensionInstallForcelist","ExtensionInstallAllowlist","ExtensionInstallSources")) {
        $path = "HKLM:\SOFTWARE\Policies\$key\$leaf"
        if (Test-Path $path) { Remove-Item -Path $path -Recurse -Force; Write-Host "Removed $path" }
    }
}

$installDir = "C:\ProgramData\Zimunim"
if (Test-Path $installDir) { Remove-Item $installDir -Recurse -Force; Write-Host "Removed $installDir" }

Write-Host "Uninstalled. Restart browsers." -ForegroundColor Yellow
