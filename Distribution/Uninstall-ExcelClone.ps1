# ExcelClone — Windows Uninstaller Script
param ()

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Uninstalling ExcelClone Shortcuts & Associations        " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Desktop Shortcut
$desktopShortcut = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)) "ExcelClone.lnk"
if (Test-Path $desktopShortcut) {
    Remove-Item $desktopShortcut -Force
    Write-Host "[OK] Desktop shortcut removed." -ForegroundColor Green
}

# 2. Start Menu Shortcut
$startMenuAppDir = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::Programs)) "ExcelClone"
if (Test-Path $startMenuAppDir) {
    Remove-Item $startMenuAppDir -Recurse -Force
    Write-Host "[OK] Start Menu shortcut folder removed." -ForegroundColor Green
}

# 3. File Associations in HKCU
try {
    $progId = "ExcelClone.Spreadsheet"
    $hkcuClasses = "HKCU:\Software\Classes"

    if (Test-Path "$hkcuClasses\$progId") {
        Remove-Item "$hkcuClasses\$progId" -Recurse -Force -ErrorAction SilentlyContinue
    }

    $extensions = @(".xlsx", ".xls", ".xlsm", ".csv")
    foreach ($ext in $extensions) {
        if (Test-Path "$hkcuClasses\$ext\OpenWithProgids") {
            Remove-ItemProperty -Path "$hkcuClasses\$ext\OpenWithProgids" -Name $progId -ErrorAction SilentlyContinue
        }
    }
    Write-Host "[OK] File associations unregistered." -ForegroundColor Green
}
catch {
    Write-Host "[NOTICE] File association unregistration completed." -ForegroundColor Yellow
}

Write-Host "`nExcelClone shortcuts and associations removed successfully." -ForegroundColor Green
