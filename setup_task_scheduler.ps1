# Script PowerShell untuk daftarkan sync_to_excel.py ke Windows Task Scheduler
# Jalankan sebagai Administrator: klik kanan → "Run as Administrator"

$pythonPath  = (Get-Command python).Source
$scriptPath  = "C:\Claude\Project\Financial\sync_to_excel.py"
$workingDir  = "C:\Claude\Project\Financial"
$taskName    = "FinancialBot_SyncExcel"

# Hapus task lama jika ada
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

# Buat action
$action = New-ScheduledTaskAction `
    -Execute $pythonPath `
    -Argument "`"$scriptPath`"" `
    -WorkingDirectory $workingDir

# Trigger 1: Saat Windows startup
$triggerBoot = New-ScheduledTaskTrigger -AtStartup

# Trigger 2: Setiap 2 menit
$triggerRepeat = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 1) `
    -Once -At (Get-Date)

# Setting: jalankan walau tidak ada user login, highest privilege
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 1) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)

$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `
    -RunLevel Highest

# Daftarkan task
Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $triggerBoot, $triggerRepeat `
    -Settings $settings `
    -Principal $principal `
    -Description "Sync transaksi dari Telegram bot ke Financial Management.xlsx" `
    -Force

Write-Host ""
Write-Host "Task Scheduler berhasil dibuat: $taskName" -ForegroundColor Green
Write-Host "Sync akan berjalan:"
Write-Host "  - Saat Windows startup"
Write-Host "  - Setiap 2 menit selama PC menyala"
Write-Host ""
Write-Host "Cek log di: $workingDir\sync.log"
