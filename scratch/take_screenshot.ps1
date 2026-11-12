Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$proc = Start-Process -FilePath "C:\Users\it\Documents\4D PO\clone exc\EFM_BOM_App\bin\Release\net10.0-windows\EFM_BOM_App.exe" -PassThru
Start-Sleep -Seconds 4

$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

$artifactDir = "C:\Users\it\.gemini\antigravity\brain\69026510-822f-4af9-87e1-7017bb7376eb"
$screenshotPath = Join-Path $artifactDir "screenshot_excelclone_mvp.png"
$bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)

$graphics.Dispose()
$bitmap.Dispose()

# Keep process running for 2 seconds then stop
Start-Sleep -Seconds 2
Stop-Process -Id $proc.Id -Force

Write-Host "Screenshot saved to $screenshotPath"

