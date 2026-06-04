$ErrorActionPreference = "Stop"

$repoRoot   = (Resolve-Path "$PSScriptRoot\..").Path
$distDir    = Join-Path $repoRoot "dist"
$policyDir  = Join-Path $distDir "policy"
$crxFile    = Join-Path $distDir "zimunim.crx"
$manifestFile = Join-Path $distDir "extension\manifest.json"

if (-not (Test-Path $crxFile))      { throw "Missing $crxFile - run build-crx.ps1 first" }
if (-not (Test-Path $manifestFile)) { throw "Missing $manifestFile - run build-crx.ps1 first" }

$installDir = "C:\ProgramData\Zimunim"

function Read-Varint([byte[]]$d, [ref]$p) {
    $r = [uint64]0; $s = 0
    while ($true) {
        $b = $d[$p.Value]; $p.Value++
        $r = $r -bor (([uint64]($b -band 0x7F)) -shl $s)
        if (($b -band 0x80) -eq 0) { return $r }
        $s += 7
    }
}

function Get-CrxExtensionId([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $headerSize = [BitConverter]::ToUInt32($bytes, 8)
    $headerEnd = 12 + $headerSize
    $pos = 12
    $pubKey = $null
    while ($pos -lt $headerEnd -and -not $pubKey) {
        $tag = $bytes[$pos]; $pos++
        $pr  = [ref]$pos
        $len = [int](Read-Varint $bytes $pr); $pos = $pr.Value
        if ($tag -eq 0x12) {
            $end = $pos + $len
            while ($pos -lt $end -and -not $pubKey) {
                $itag = $bytes[$pos]; $pos++
                $pr   = [ref]$pos
                $ilen = [int](Read-Varint $bytes $pr); $pos = $pr.Value
                if ($itag -eq 0x0A) { $pubKey = $bytes[$pos..($pos + $ilen - 1)] }
                $pos += $ilen
            }
            $pos = $end
        } else { $pos += $len }
    }
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($pubKey)
    $hex  = -join ($hash | ForEach-Object { $_.ToString("x2") })
    $map  = @{"0"="a";"1"="b";"2"="c";"3"="d";"4"="e";"5"="f";"6"="g";"7"="h";
              "8"="i";"9"="j";"a"="k";"b"="l";"c"="m";"d"="n";"e"="o";"f"="p"}
    return -join ($hex.Substring(0,32).ToCharArray() | ForEach-Object { $map["$_"] })
}

$manifest = Get-Content $manifestFile -Raw | ConvertFrom-Json
$version  = $manifest.version
$extId    = Get-CrxExtensionId $crxFile

Write-Host "Extension ID : $extId" -ForegroundColor Green
Write-Host "Version      : $version" -ForegroundColor Green
Write-Host "Install dir  : $installDir" -ForegroundColor Green

if (Test-Path $policyDir) { Remove-Item $policyDir -Recurse -Force }
New-Item -ItemType Directory -Path $policyDir | Out-Null

$crxUrl = "file:///" + ($installDir -replace '\\','/') + "/zimunim.crx"
$xmlUrl = "file:///" + ($installDir -replace '\\','/') + "/update.xml"

@"
<?xml version='1.0' encoding='UTF-8'?>
<gupdate xmlns='http://www.google.com/update2/response' protocol='2.0'>
  <app appid='$extId'>
    <updatecheck codebase='$crxUrl' version='$version' />
  </app>
</gupdate>
"@ | Set-Content -Path (Join-Path $policyDir "update.xml") -Encoding UTF8

$browsers = @(
    @{ name = "Chrome"; key  = "Google\Chrome" },
    @{ name = "Brave";  key  = "BraveSoftware\Brave" },
    @{ name = "Edge";   key  = "Microsoft\Edge" }
)

function Build-RegBody([string]$keyPath, [string]$extId, [string]$xmlUrl) {
@"
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\$keyPath\ExtensionInstallForcelist]
"1"="$extId;$xmlUrl"

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\$keyPath\ExtensionInstallAllowlist]
"1"="$extId"

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\$keyPath\ExtensionInstallSources]
"1"="file:///*"
"@
}

function Build-UninstallBody([string]$keyPath) {
@"
[-HKEY_LOCAL_MACHINE\SOFTWARE\Policies\$keyPath\ExtensionInstallForcelist]
[-HKEY_LOCAL_MACHINE\SOFTWARE\Policies\$keyPath\ExtensionInstallAllowlist]
[-HKEY_LOCAL_MACHINE\SOFTWARE\Policies\$keyPath\ExtensionInstallSources]
"@
}

$allInstall = @("Windows Registry Editor Version 5.00", "")
$allRemove  = @("Windows Registry Editor Version 5.00", "")

foreach ($b in $browsers) {
    $body = Build-RegBody $b.key $extId $xmlUrl
    $undo = Build-UninstallBody $b.key

    $singleInstall = @("Windows Registry Editor Version 5.00", "", $body) -join "`r`n"
    $singleRemove  = @("Windows Registry Editor Version 5.00", "", $undo) -join "`r`n"

    Set-Content -Path (Join-Path $policyDir ("install-{0}.reg" -f $b.name.ToLower())) -Value $singleInstall -Encoding Unicode
    Set-Content -Path (Join-Path $policyDir ("uninstall-{0}.reg" -f $b.name.ToLower())) -Value $singleRemove -Encoding Unicode

    $allInstall += $body, ""
    $allRemove  += $undo, ""
}

Set-Content -Path (Join-Path $policyDir "install-all.reg")   -Value ($allInstall -join "`r`n") -Encoding Unicode
Set-Content -Path (Join-Path $policyDir "uninstall-all.reg") -Value ($allRemove  -join "`r`n") -Encoding Unicode

$installerScript = @"
#Requires -RunAsAdministrator
`$ErrorActionPreference = "Stop"

`$here       = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$installDir = "$installDir"
`$extId      = "$extId"
`$version    = "$version"

if (-not (Test-Path `$installDir)) { New-Item -ItemType Directory -Path `$installDir | Out-Null }
Copy-Item -Path (Join-Path `$here "..\zimunim.crx") -Destination `$installDir -Force
Copy-Item -Path (Join-Path `$here "update.xml")     -Destination `$installDir -Force

`$browsers = @("Google\Chrome","BraveSoftware\Brave","Microsoft\Edge")
foreach (`$key in `$browsers) {
    `$base = "HKLM:\SOFTWARE\Policies\`$key"
    foreach (`$leaf in @("ExtensionInstallForcelist","ExtensionInstallAllowlist","ExtensionInstallSources")) {
        if (-not (Test-Path "`$base\`$leaf")) { New-Item -Path "`$base\`$leaf" -Force | Out-Null }
    }
    Set-ItemProperty -Path "`$base\ExtensionInstallForcelist" -Name "1" -Value "`$extId;file:///`$(`$installDir -replace '\\','/')/update.xml"
    Set-ItemProperty -Path "`$base\ExtensionInstallAllowlist" -Name "1" -Value `$extId
    Set-ItemProperty -Path "`$base\ExtensionInstallSources"   -Name "1" -Value "file:///*"
    Write-Host "Configured policy for `$key" -ForegroundColor Green
}

Write-Host ""
Write-Host "Extension `$extId v`$version installed via policy." -ForegroundColor Green
Write-Host "Restart Chrome/Brave/Edge to load the extension." -ForegroundColor Yellow
"@

$uninstallerScript = @"
#Requires -RunAsAdministrator
`$ErrorActionPreference = "Continue"

`$browsers = @("Google\Chrome","BraveSoftware\Brave","Microsoft\Edge")
foreach (`$key in `$browsers) {
    foreach (`$leaf in @("ExtensionInstallForcelist","ExtensionInstallAllowlist","ExtensionInstallSources")) {
        `$path = "HKLM:\SOFTWARE\Policies\`$key\`$leaf"
        if (Test-Path `$path) { Remove-Item -Path `$path -Recurse -Force; Write-Host "Removed `$path" }
    }
}

`$installDir = "$installDir"
if (Test-Path `$installDir) { Remove-Item `$installDir -Recurse -Force; Write-Host "Removed `$installDir" }

Write-Host "Uninstalled. Restart browsers." -ForegroundColor Yellow
"@

Set-Content -Path (Join-Path $policyDir "install.ps1")   -Value $installerScript   -Encoding UTF8
Set-Content -Path (Join-Path $policyDir "uninstall.ps1") -Value $uninstallerScript -Encoding UTF8

$readme = @"
=== Zimunim Extension - Policy Deployment ===

Extension ID: $extId
Version     : $version
Install path: $installDir

----- Easiest install (recommended) -----
1. Right-click install.ps1 -> Run with PowerShell as Administrator.
   (If blocked: open elevated PowerShell, then:
    Set-ExecutionPolicy -Scope Process Bypass; .\install.ps1)
2. Restart Chrome / Brave / Edge.
3. Verify in chrome://extensions - the extension appears with
   "Installed by enterprise policy".

----- Manual install with REG files -----
1. Copy zimunim.crx and update.xml to: $installDir
2. Double-click install-all.reg (or install-chrome.reg / brave / edge for a
   single browser) and confirm.
3. Restart the browser.

If the CRX or update.xml are placed somewhere other than $installDir,
edit the file:/// paths inside update.xml and the .reg file accordingly.

----- Uninstall -----
- Run uninstall.ps1 as Administrator, OR
- Double-click uninstall-all.reg.

----- Files in this folder -----
  install.ps1           - smart installer (admin)
  uninstall.ps1         - smart uninstaller (admin)
  update.xml            - update manifest (referenced by the policy)
  install-all.reg       - REG for Chrome + Brave + Edge
  install-chrome.reg    - REG for Chrome only
  install-brave.reg     - REG for Brave only
  install-edge.reg      - REG for Edge only
  uninstall-all.reg     - removes policy from all three
  uninstall-<browser>.reg - removes policy from a single browser

The zimunim.crx file lives one folder up (dist\zimunim.crx). Keep it
together with this folder when distributing.
"@

Set-Content -Path (Join-Path $policyDir "README.txt") -Value $readme -Encoding UTF8

Write-Host ""
Write-Host "==== Policy files in $policyDir ====" -ForegroundColor Green
Get-ChildItem $policyDir | Select-Object Name, @{n="Size";e={"{0:N0} B" -f $_.Length}} | Format-Table -AutoSize
