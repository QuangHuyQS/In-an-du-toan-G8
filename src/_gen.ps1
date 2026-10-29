# Sinh ma nguon VBA (ANSI, CRLF) tu src\ -> gen\   (ban PowerShell cua _gen.py, dung khi may khong co Python)
# Chuoi "..." co ky tu tieng Viet -> U("...\uXXXX...") vi VBE chi doc ANSI.
$ErrorActionPreference = 'Stop'
$SRC = $PSScriptRoot
$GEN = Join-Path (Split-Path $SRC -Parent) 'gen'
New-Item -ItemType Directory -Force $GEN | Out-Null

$CLS_HDR = "VERSION 1.0 CLASS`nBEGIN`n  MultiUse = -1  'True`nEND`nAttribute VB_Name = ""{0}""`nAttribute VB_GlobalNameSpace = False`nAttribute VB_Creatable = False`nAttribute VB_PredeclaredId = False`nAttribute VB_Exposed = False`nOption Explicit`n"

function Esc([string]$s) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        if ([int]$ch -lt 128) { [void]$sb.Append($ch) } else { [void]$sb.Append(('\u{0:X4}' -f [int]$ch)) }
    }
    $sb.ToString()
}
function HasNonAscii([string]$s) { foreach ($ch in $s.ToCharArray()) { if ([int]$ch -gt 127) { return $true } }; $false }

function ConvertLine([string]$line) {
    $out = New-Object System.Text.StringBuilder
    $i = 0; $n = $line.Length
    while ($i -lt $n) {
        $c = $line[$i]
        if ($c -eq "'") {
            $rest = $line.Substring($i)
            if (HasNonAscii $rest) { throw "non-ASCII in comment: $line" }
            [void]$out.Append($rest); break
        }
        if ($c -eq '"') {
            $j = $i + 1
            while ($true) {
                $k = $line.IndexOf('"', $j)
                if ($k -lt 0) { throw "unterminated string: $line" }
                if ($k + 1 -lt $n -and $line[$k + 1] -eq '"') { $j = $k + 2; continue }
                break
            }
            $lit = $line.Substring($i, $k - $i + 1)
            if (HasNonAscii $lit) { [void]$out.Append('U(' + (Esc $lit) + ')') } else { [void]$out.Append($lit) }
            $i = $k + 1; continue
        }
        [void]$out.Append($c); $i++
    }
    $out.ToString()
}

function Gen([string]$name, [string]$outname = $null, [string]$cls = $null) {
    $txt = [IO.File]::ReadAllText((Join-Path $SRC $name), [Text.Encoding]::UTF8)
    $lines = $txt.Replace("`r`n", "`n").Split("`n") | ForEach-Object { ConvertLine $_ }
    $body = $lines -join "`n"
    if ($cls) { $body = ($CLS_HDR -f $cls) + $body }
    if (HasNonAscii $body) { throw "gen $name con ky tu non-ASCII" }
    if (-not $outname) { $outname = $name }
    $body = $body.Replace("`n", "`r`n")
    [IO.File]::WriteAllText((Join-Path $GEN $outname), $body, [Text.Encoding]::ASCII)
    Write-Host "gen $outname"
}

Gen 'inanDTG8_V6.bas'
Gen 'modInAnMain.bas'
Gen 'BuildInAnForm.bas'
Gen 'frmInDuToan_code.txt'
Gen 'clsInAnBtn.cls' -cls 'clsInAnBtn'
