# BUILD_V6.ps1 - dung file tong FILE_TONG_IN_G8_V6.xlsm tu ma nguon V6\src
# Chay:  powershell -ExecutionPolicy Bypass -File BUILD_V6.ps1
# Can: Excel + "Trust access to the VBA project object model" + Python (py) de sinh gen\.
# Chay tren 1 phien Excel RIENG, khong dung toi file ban dang mo.
param(
    [string]$Template = (Join-Path $PSScriptRoot '..\FILE_TONG_IN_G8_khung.xlsm'),
    [string]$Out      = (Join-Path $PSScriptRoot 'FILE_TONG_IN_G8_V6.2.xlsm'),
    # -KeepForm: giu nguyen form frmInDuToan + clsInAnBtn cua file khung (trinh thiet ke form cua VBE hay loi 75 khi dung bang code)
    [switch]$KeepForm
)
$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'src'
$gen = Join-Path $PSScriptRoot 'gen'

# Khong co file khung -> dung ban V6.1 da dung truoc do lam khung (code/form deu bi thay lai o duoi)
if (-not (Test-Path $Template)) { $Template = Join-Path $PSScriptRoot 'FILE_TONG_IN_G8_V6.1.xlsm' }

# Sinh gen\ : uu tien Python that; khong co thi dung ban PowerShell
$pyOk = $false
try { $v = & py -I -c "print(1)" 2>$null; if ($LASTEXITCODE -eq 0 -and "$v" -eq '1') { $pyOk = $true } } catch {}
if ($pyOk) {
    & py -I (Join-Path $src '_gen.py')
    if ($LASTEXITCODE -ne 0) { throw 'Sinh ma gen\ that bai' }
} else {
    & (Join-Path $src '_gen.ps1')
}

$Template = (Resolve-Path $Template).Path
if (Test-Path $Out) { Remove-Item $Out -Force }
Copy-Item $Template $Out

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $wb = $xl.Workbooks.Open($Out)
    $vbp = $wb.VBProject

    $rm = @('BuildFormsG8','modFormBridge','frmToolInDuToan','frmIn2HangMuc','KHUNG_G8_Launcher','inanDTG8_V5','inanDTG8_V6','modInAnMain','BuildInAnForm')
    if (-not $KeepForm) { $rm += 'frmInDuToan','clsInAnBtn' }
    foreach ($n in $rm) {
        try { $c = $vbp.VBComponents.Item($n); $vbp.VBComponents.Remove($c); Write-Host "removed $n" } catch {}
    }
    foreach ($f in 'inanDTG8_V6.bas','modInAnMain.bas','BuildInAnForm.bas') {
        [void]$vbp.VBComponents.Import((Join-Path $gen $f)); Write-Host "imported $f"
    }

    if (-not $KeepForm) {
        $xl.Run('BuildInAnForm_Run', (Join-Path $gen 'frmInDuToan_code.txt'), (Join-Path $gen 'clsInAnBtn.cls'))
        Write-Host 'form built'
    } else {
        # Giu giao dien form, chi thay code (khong dung trinh thiet ke -> khong loi 75)
        $cm = $vbp.VBComponents.Item('frmInDuToan').CodeModule
        $txt = [IO.File]::ReadAllLines((Join-Path $gen 'frmInDuToan_code.txt'), [Text.Encoding]::Default) |
            Where-Object { $_ -notmatch '^\s*(Attribute |VERSION |BEGIN|END$)' }
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString(($txt -join "`r`n"))
        Write-Host 'form code replaced'
    }
    $xl.Run('BuildInAnForm_Sheet')
    Write-Host 'sheet updated'

    $vbp.VBComponents.Remove($vbp.VBComponents.Item('BuildInAnForm'))

    # Xuat form/class ra gen\ de tham khao / import tay
    if (-not $KeepForm) { $vbp.VBComponents.Item('frmInDuToan').Export((Join-Path $gen 'frmInDuToan.frm')) }

    $wb.Save()
    $wb.Close($false)
    Write-Host "OK -> $Out"
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
