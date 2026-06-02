# Deploy ke Hugging Face Spaces

Bot akan jalan 24/7 di cloud. Laptop nggak perlu hidup — saat laptop hidup,
`sync_to_excel.py` (Task Scheduler) otomatis pull semua transaksi pending dan
tulis ke Excel + push ke GitHub.

## Step 1 — Daftar Hugging Face

1. Buka https://huggingface.co/join
2. Daftar pakai email (atau Google/GitHub login)
3. Verifikasi email
4. Pilih username (catat — ini muncul di URL Space)

## Step 2 — Buat User Access Token (untuk push code)

1. Buka https://huggingface.co/settings/tokens
2. Klik **New token**
3. Name: `financial-bot-deploy`
4. Type: **Write** (penting — read-only tidak bisa push)
5. Klik **Create token**
6. **Copy token** (`hf_xxxxxxxxxxxx`) — simpan di tempat aman, hanya muncul sekali

## Step 3 — Buat Space baru

1. Buka https://huggingface.co/new-space
2. Owner: username kamu
3. Space name: `financial-bot`
4. License: `MIT` (atau pilihan lain)
5. SDK: **Docker** ← penting
6. Hardware: **CPU basic - Free** (cukup)
7. Visibility: **Private** (disarankan karena ada token bot)
8. Klik **Create Space**

Setelah dibuat, URL akan jadi: `https://huggingface.co/spaces/<username>/financial-bot`

## Step 4 — Set Secrets di Space

1. Di halaman Space → tab **Settings**
2. Scroll ke **Variables and secrets** → **New secret**
3. Tambahkan satu-per-satu (klik **Save** setelah tiap secret):

| Name | Value |
|------|-------|
| `TELEGRAM_BOT_TOKEN` | Token bot dari @BotFather (yang sekarang ada di `.env` lokal) |
| `ALLOWED_USER_ID` | `8353578388` (user ID kamu, dari `.env`) |
| `SYNC_SECRET` | String acak — bisa pakai yang dari `.env` lokal, ATAU generate baru |

> Note: `CLOUD_MODE=1`, `PORT=7860`, `DB_PATH=/data/transactions.db` sudah ter-set
> via Dockerfile. Tidak perlu set di Settings.

## Step 5 — Enable Persistent Storage (penting!)

1. Settings → scroll ke **Persistent Storage**
2. Klik **Enable** → pilih plan terkecil (free tier biasanya ada — kalau tidak,
   skip step ini, risiko: SQLite reset saat restart, tapi sync 1-menit dari laptop
   nge-catch sebelum sempat hilang)
3. Mount point: `/data` (default)

## Step 6 — Push code ke HF Space

Di terminal laptop, dari folder Financial:

```bash
cd "C:\Claude\Project\Financial"

# Tambahkan HF Space sebagai remote git (ganti USERNAME dan TOKEN)
git remote add hf https://USERNAME:hf_xxxxxxxxxxxx@huggingface.co/spaces/USERNAME/financial-bot

# Push code ke HF
git push hf master
```

HF akan auto-build Docker image (5-10 menit). Pantau di tab **App** di halaman Space.

## Step 7 — Verifikasi bot di cloud

1. Buka `https://<username>-financial-bot.hf.space/health` di browser
   → harus muncul `{"status":"ok"}`
2. Kirim `/help` ke bot di Telegram → harus dapat balasan
3. Cek logs di Space tab **Logs**: harus muncul `Bot aktif — sheet bulan ini: ...`

## Step 8 — Cutover laptop ke mode "client only"

Setelah bot cloud confirmed jalan, switch laptop dari hosting bot → cuma sync:

```powershell
# 1. Stop bot lokal & unregister task-nya
Unregister-ScheduledTask -TaskName 'FinancialBot_Telegram' -Confirm:$false
Stop-Process -Name pythonw -Force -EA SilentlyContinue

# 2. Update .env: ganti RAILWAY_URL ke URL cloud kamu
# Buka file .env, ubah baris:
#   RAILWAY_URL=http://localhost:8080
# Jadi:
#   RAILWAY_URL=https://<username>-financial-bot.hf.space

# 3. Register sync task (jalan AtStartup + tiap 1 menit)
& "C:\Claude\Project\Financial\setup_task_scheduler.ps1"

# 4. Restart dashboard task supaya pakai data terbaru
Restart-ScheduledTask -TaskName 'FinancialBot_Dashboard'
```

## Step 9 — Test end-to-end

1. **Kirim tx ke bot via Telegram**: `beli kopi 50nt`
2. **Bot di cloud** terima, simpan ke SQLite cloud
3. **Tunggu max 1 menit** → sync_to_excel.py (Task Scheduler) pull dari cloud
4. **Cek Excel**: tx baru muncul di sheet bulan ini
5. **Cek dashboard** di `localhost:5050`: bar di hari ini bertambah

Done. Sekarang:
- **Laptop mati / restart** → bot tetap terima tx, simpan di cloud SQLite
- **Laptop hidup** → AtStartup trigger → sync langsung jalan → semua tx pending masuk Excel + dashboard update

## Troubleshooting

- **Build gagal**: cek Space tab **Logs** → biasanya error di Dockerfile / requirements.txt
- **Bot 409 Conflict di log**: ada instance bot lain (di laptop atau Space lain) masih jalan dengan token sama → stop yang lain
- **`/health` 404**: belum selesai build atau crash → cek logs
- **Sync gagal**: cek `RAILWAY_URL` di `.env` benar (https://, tanpa trailing slash), dan SYNC_SECRET sama di Space + .env
