# Build standalone Daily Data.exe and auto-send to Desktop
$ErrorActionPreference = "Stop"

Write-Host "1. Menghentikan proses Daily Data yang sedang berjalan..." -ForegroundColor Cyan
Get-Process -Name "*Daily Data*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

Write-Host "2. Membangun executable standalone dengan PyInstaller..." -ForegroundColor Cyan
& "App\.venv\Scripts\pyinstaller.exe" --noconfirm --onefile --windowed --name "Daily Data" --distpath "." --workpath "build" App/gui.py

Write-Host "3. Membersihkan file Daily Data lama di Desktop..." -ForegroundColor Cyan
$desktopPath = [Environment]::GetFolderPath('Desktop')
if (-not $desktopPath -or -not (Test-Path $desktopPath)) {
    $desktopPath = "C:\Users\Akmal\OneDrive\Desktop"
}

# Hapus SEMUA file Daily Data*.exe yang ada di desktop terlebih dahulu
Get-ChildItem -Path $desktopPath -Filter "*Daily Data*.exe" -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "   -> Menghapus file lama: $($_.FullName)" -ForegroundColor Yellow
    Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
}

Write-Host "4. Menempatkan Daily Data.exe terbaru ke Desktop..." -ForegroundColor Cyan
$targetDesktop = Join-Path $desktopPath "Daily Data.exe"
Copy-Item -Path "Daily Data.exe" -Destination $targetDesktop -Force

Write-Host "5. Membersihkan direktori build sementara..." -ForegroundColor Cyan
Remove-Item -Recurse -Force "build", "dist" -ErrorAction SilentlyContinue

Write-Host "`n[SUKSES] Daily Data.exe terbaru berhasil dibungkus dan tersedia di Desktop (bersih tanpa duplikat):" -ForegroundColor Green
Write-Host " -> $targetDesktop" -ForegroundColor Green
