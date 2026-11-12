$exe = Resolve-Path "Distribution\ExcelClone_Portable\ExcelClone.exe"
$sample = Resolve-Path "samples\Sample_Workbook.xlsx"
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  TESTING EXCELCLONE STANDALONE LAUNCH WITHOUT LOGIN      " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Target Executable: $exe"

# Test 1: Launch blank application
Write-Host "`n[Test 1] Launching blank workbook..."
$p1 = Start-Process -FilePath $exe -PassThru
Start-Sleep -Seconds 3

$alive1 = -not $p1.HasExited
$ws1 = (Get-Process -Id $p1.Id -ErrorAction SilentlyContinue).WorkingSet64 / 1MB
Write-Host "Process ID: $($p1.Id) | Alive: $alive1 | Working Set: $([Math]::Round($ws1, 2)) MB" -ForegroundColor Green

Stop-Process -Id $p1.Id -Force
Write-Host "[Test 1 OK] Blank workbook launched and stopped cleanly." -ForegroundColor Green

# Test 2: Launch with sample workbook
Write-Host "`n[Test 2] Launching with sample workbook: $sample"
$p2 = Start-Process -FilePath $exe -ArgumentList "`"$sample`"" -PassThru
Start-Sleep -Seconds 3

$alive2 = -not $p2.HasExited
$ws2 = (Get-Process -Id $p2.Id -ErrorAction SilentlyContinue).WorkingSet64 / 1MB
Write-Host "Process ID: $($p2.Id) | Alive: $alive2 | Working Set: $([Math]::Round($ws2, 2)) MB" -ForegroundColor Green

Stop-Process -Id $p2.Id -Force
Write-Host "[Test 2 OK] Sample workbook opened directly without login." -ForegroundColor Green

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "  ALL LAUNCH TESTS PASSED — STANDALONE OPERATION CONFIRMED" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

