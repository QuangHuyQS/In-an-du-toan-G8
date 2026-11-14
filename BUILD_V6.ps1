# BUILD_V6.ps1 - nap ma nguon VBA (src\) vao file BO MAY FILE_TONG_IN_G8_V*.xlsm cho phan mem InAnG8.exe
# Chay:  powershell -ExecutionPolicy Bypass -File BUILD_V6.ps1 -Out <xlsm moi> [-Template <xlsm khung>]
#   Khung mac dinh = FILE_TONG_IN_G8_V7.1.xlsm canh script. -Out phai KHAC khung; build xong kiem roi chep de len.
# Can: Excel + "Trust access to the VBA project object model" + Python (py) de sinh gen\.
# Tu V7.1 file xlsm chi la bo may: KHONG con form VBA (nen xanh, nut cam) - giao dien la InAnG8.exe.
# Chay tren 1 phien Excel RIENG, khong dung toi file ban dang mo.
param(
    [string]$Template = '',
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Template) { $Template = Join-Path $here 'FILE_TONG_IN_G8_V7.1.xlsm' }
$src = Join-Path $here 'src'
$gen = Join-Path $here 'gen'

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
$OutFull = [IO.Path]::GetFullPath($Out)
if ($OutFull -ieq $Template) { throw '-Out phai khac -Template (build ra file khac roi chep de len)' }
if (Test-Path $OutFull) { Remove-Item $OutFull -Force }
Copy-Item $Template $OutFull

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $wb = $xl.Workbooks.Open($OutFull)
    $vbp = $wb.VBProject

    # bo module cu + form VBA cu (frmInDuToan, clsInAnBtn) neu file khung con
    $rm = @('BuildFormsG8','modFormBridge','frmToolInDuToan','frmIn2HangMuc','KHUNG_G8_Launcher','inanDTG8_V5',
            'inanDTG8_V6','modInAnMain','BuildInAnForm','frmInDuToan','clsInAnBtn')
    foreach ($n in $rm) {
        try { $c = $vbp.VBComponents.Item($n); $vbp.VBComponents.Remove($c); Write-Host "removed $n" } catch {}
    }
    foreach ($f in 'inanDTG8_V6.bas','modInAnMain.bas','BuildInAnForm.bas') {
        [void]$vbp.VBComponents.Import((Join-Path $gen $f)); Write-Host "imported $f"
    }

    # sheet cau hinh "IN THEO FILE TONG" (engine doc C12/C13): hien lai de dung lai, BuildInAnForm_Engine an di
    try { $wb.Worksheets.Item('IN THEO FILE TONG').Visible = -1 } catch {}
    $xl.Run('BuildInAnForm_Sheet')
    $xl.Run('BuildInAnForm_Engine')
    Write-Host 'sheet: chi con trang huong dan mo InAnG8.exe'

    $vbp.VBComponents.Remove($vbp.VBComponents.Item('BuildInAnForm'))

    $wb.Save()
    $wb.Close($false)
    Write-Host "OK -> $OutFull"
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
