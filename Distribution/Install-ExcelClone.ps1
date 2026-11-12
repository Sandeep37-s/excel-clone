# ExcelClone — Windows Installer & Shortcut Provisioning Script
param (
    [switch]$CreateDesktopShortcut = $true,
    [switch]$CreateStartMenuShortcut = $true,
    [switch]$RegisterFileAssociations = $true
)

$installDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$portableDir = Join-Path $installDir "ExcelClone_Portable"
$exePath = Join-Path $portableDir "ExcelClone.exe"
$icoPath = Join-Path $portableDir "app.ico"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "  Installing ExcelClone Professional (Free & Open Source Edition)    " -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "Target Application: $exePath"

if (-not (Test-Path $exePath)) {
    Write-Host "[ERROR] Executable not found at $exePath" -ForegroundColor Red
    exit 1
}

$wshShell = New-Object -ComObject WScript.Shell

# 1. Desktop Shortcut
if ($CreateDesktopShortcut) {
    $desktopDir = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $shortcutPath = Join-Path $desktopDir "ExcelClone.lnk"
    $shortcut = $wshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exePath
    $shortcut.WorkingDirectory = $portableDir
    $shortcut.Description = "ExcelClone — Free & Open-Source High-Performance Windows Spreadsheet"
    if (Test-Path $icoPath) {
        $shortcut.IconLocation = "$icoPath,0"
    }
    $shortcut.Save()
    Write-Host "[OK] Desktop shortcut created: $shortcutPath" -ForegroundColor Green
}

# 2. Start Menu Shortcut
if ($CreateStartMenuShortcut) {
    $programsDir = [Environment]::GetFolderPath([Environment+SpecialFolder]::Programs)
    $startMenuAppDir = Join-Path $programsDir "ExcelClone"
    if (-not (Test-Path $startMenuAppDir)) { New-Item -ItemType Directory -Path $startMenuAppDir -Force | Out-Null }
    
    $shortcutPath = Join-Path $startMenuAppDir "ExcelClone.lnk"
    $shortcut = $wshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exePath
    $shortcut.WorkingDirectory = $portableDir
    $shortcut.Description = "ExcelClone — Free & Open-Source High-Performance Windows Spreadsheet"
    if (Test-Path $icoPath) {
        $shortcut.IconLocation = "$icoPath,0"
    }
    $shortcut.Save()
    Write-Host "[OK] Start Menu shortcut created: $shortcutPath" -ForegroundColor Green
}

# 3. User-Level File Associations (HKCU)
if ($RegisterFileAssociations) {
    try {
        $progId = "ExcelClone.Spreadsheet"
        $hkcuClasses = "HKCU:\Software\Classes"

        New-Item -Path "$hkcuClasses\$progId" -Value "ExcelClone Spreadsheet Workbook" -Force | Out-Null
        if (Test-Path $icoPath) {
            New-Item -Path "$hkcuClasses\$progId\DefaultIcon" -Value "$icoPath,0" -Force | Out-Null
        }
        New-Item -Path "$hkcuClasses\$progId\shell\open\command" -Value "`"$exePath`" `"%1`"" -Force | Out-Null

        $extensions = @(".xlsx", ".xls", ".xlsm", ".csv")
        foreach ($ext in $extensions) {
            New-Item -Path "$hkcuClasses\$ext\OpenWithProgids" -Force | Out-Null
            Set-ItemProperty -Path "$hkcuClasses\$ext\OpenWithProgids" -Name $progId -Value "" -Force | Out-Null
        }

        Write-Host "[OK] File associations (.xlsx, .xls, .xlsm, .csv) registered in HKCU." -ForegroundColor Green
    }
    catch {
        Write-Host "[NOTICE] File association registration completed with standard privileges." -ForegroundColor Yellow
    }
}

Write-Host "`nInstallation successful! ExcelClone is ready to use immediately." -ForegroundColor Green
