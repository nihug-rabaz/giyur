$ErrorActionPreference = "Stop"

$crxPath = Join-Path (Resolve-Path "$PSScriptRoot\..").Path "dist\zimunim.crx"
if (-not (Test-Path $crxPath)) { throw "CRX not found at $crxPath - run build-crx.ps1 first" }

function Read-Varint {
    param([byte[]]$data, [ref]$pos)
    $result = [uint64]0
    $shift  = 0
    while ($true) {
        $b = $data[$pos.Value]
        $pos.Value++
        $result = $result -bor (([uint64]($b -band 0x7F)) -shl $shift)
        if (($b -band 0x80) -eq 0) { return $result }
        $shift += 7
    }
}

$bytes = [System.IO.File]::ReadAllBytes($crxPath)
if ([System.Text.Encoding]::ASCII.GetString($bytes, 0, 4) -ne "Cr24") { throw "Not a CRX file" }

$headerSize = [BitConverter]::ToUInt32($bytes, 8)
$headerEnd  = 12 + $headerSize
$pos        = 12
$pubKey     = $null

while ($pos -lt $headerEnd -and -not $pubKey) {
    $tag = $bytes[$pos]; $pos++
    $posRef = [ref]$pos
    $len    = [int](Read-Varint $bytes $posRef)
    $pos    = $posRef.Value
    if ($tag -eq 0x12) {
        $innerEnd = $pos + $len
        while ($pos -lt $innerEnd -and -not $pubKey) {
            $innerTag = $bytes[$pos]; $pos++
            $posRef   = [ref]$pos
            $innerLen = [int](Read-Varint $bytes $posRef)
            $pos      = $posRef.Value
            if ($innerTag -eq 0x0A) {
                $pubKey = $bytes[$pos..($pos + $innerLen - 1)]
            }
            $pos += $innerLen
        }
        $pos = $innerEnd
    } else {
        $pos += $len
    }
}

if (-not $pubKey) { throw "Could not find public key in CRX header" }

$sha   = [System.Security.Cryptography.SHA256]::Create()
$hash  = $sha.ComputeHash($pubKey)
$hex   = -join ($hash | ForEach-Object { $_.ToString("x2") })
$first = $hex.Substring(0, 32)

$map = @{ "0"="a"; "1"="b"; "2"="c"; "3"="d"; "4"="e"; "5"="f"; "6"="g"; "7"="h";
         "8"="i"; "9"="j"; "a"="k"; "b"="l"; "c"="m"; "d"="n"; "e"="o"; "f"="p" }
$id = -join ($first.ToCharArray() | ForEach-Object { $map["$_"] })

Write-Host "Extension ID: $id" -ForegroundColor Green
Write-Host ""
Write-Host "Public key (for manifest.key, base64):" -ForegroundColor Cyan
Write-Host ([Convert]::ToBase64String($pubKey))
