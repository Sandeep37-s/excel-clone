$exe = "Distribution\ExcelClone_v1.3_RC_Portable\ExcelClone.exe"
$targetFile = (Resolve-Path "scratch\genuine_independent\Genuine_MS_Excel_Financial.xlsx").Path
Write-Host "Starting: $exe with $targetFile"

$proc = Start-Process -FilePath $exe -ArgumentList "`"$targetFile`"" -PassThru
Start-Sleep -Seconds 4

$id = $proc.Id
$isAlive = -not $proc.HasExited
$wsBytes = (Get-Process -Id $id -ErrorAction SilentlyContinue).WorkingSet64
$wsMB = [Math]::Round(($wsBytes / 1MB), 2)

Write-Host "Process ID: $id"
Write-Host "Running: $isAlive"
Write-Host "WorkingSet: $wsMB MB"

Stop-Process -Id $id -Force
Write-Host "Launch test complete."

