# Auto-register Telegram bot + Flask dashboard sebagai background services
# di Windows Task Scheduler. Tasks auto-start saat logon + auto-restart kalau crash.
#
# Jalankan: klik kanan file ini → "Run with PowerShell"
# (atau dari PowerShell: powershell -ExecutionPolicy Bypass -File setup_bot_autostart.ps1)
#
# Tidak butuh admin — semua dijalankan sebagai current user.

$ErrorActionPreference = "Stop"
$root = "C:\Claude\Project\Financial"

# Cari pythonw.exe (Python tanpa console window). Fallback ke python.exe kalau tidak ada.
$pyw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue)
if ($pyw) {
    $pythonExe = $pyw.Source
    Write-Host "Using pythonw.exe (silent): $pythonExe" -ForegroundColor Cyan
} else {
    $pythonExe = (Get-Command python.exe).Source
    Write-Host "pythonw.exe tidak ditemukan, fallback ke: $pythonExe" -ForegroundColor Yellow
}

function Register-PersistentTask {
    param(
        [string]$Name,
        [string]$Script,
        [string]$Description
    )

    # Hapus task lama kalau sudah ada
    Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction SilentlyContinue

    $action = New-ScheduledTaskAction `
        -Execute $pythonExe `
        -Argument "`"$root\$Script`"" `
        -WorkingDirectory $root

    # Trigger: saat user logon (TIDAK at startup — supaya pakai user environment & .env)
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

    # Settings untuk long-running service:
    # - StartWhenAvailable: re-trigger kalau missed
    # - RestartCount + RestartInterval: auto-restart kalau process exit/crash
    # - ExecutionTimeLimit 0 = unlimited (default ada batas waktu!)
    # - Hidden: tidak tampil di Task Scheduler UI utama
    $settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -RestartCount 999 `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit ([TimeSpan]::Zero) `
        -DontStopOnIdleEnd `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -MultipleInstances IgnoreNew

    # Principal: jalankan sebagai current user, level normal (no admin needed)
    $principal = New-ScheduledTaskPrincipal `
        -UserId $env:USERNAME `
        -LogonType Interactive `
        -RunLevel Limited

    Register-ScheduledTask `
        -TaskName $Name `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Principal $principal `
        -Description $Description `
        -Force | Out-Null

    Write-Host "  Registered: $Name" -ForegroundColor Green
}

Write-Host ""
Write-Host "Registering tasks..." -ForegroundColor White

Register-PersistentTask `
    -Name "FinancialBot_Telegram" `
    -Script "bot_cloud.py" `
    -Description "Telegram bot listener (parse natural-language transactions → SQLite → Excel)"

Register-PersistentTask `
    -Name "FinancialBot_Dashboard" `
    -Script "financial_server.py" `
    -Description "Local Flask dashboard server (http://localhost:5050)"

Write-Host ""
Write-Host "Selesai. Tasks akan auto-start saat next logon Windows." -ForegroundColor Green
Write-Host ""
Write-Host "Untuk start SEKARANG juga (kalau bot/server belum jalan):" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName 'FinancialBot_Telegram'"
Write-Host "  Start-ScheduledTask -TaskName 'FinancialBot_Dashboard'"
Write-Host ""
Write-Host "Untuk lihat status:"
Write-Host "  Get-ScheduledTask -TaskName 'FinancialBot*' | Get-ScheduledTaskInfo"
Write-Host ""
Write-Host "Untuk uninstall (kalau tidak mau auto-start lagi):"
Write-Host "  Unregister-ScheduledTask -TaskName 'FinancialBot_Telegram' -Confirm:`$false"
Write-Host "  Unregister-ScheduledTask -TaskName 'FinancialBot_Dashboard' -Confirm:`$false"
