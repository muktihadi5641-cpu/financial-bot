"""
Bot Telegram untuk Financial Management — versi CLOUD (Railway.app)
Berjalan 24/7 di server cloud. Transaksi disimpan ke SQLite.
Saat PC user nyala, sync_to_excel.py mengambil data ini dan menulis ke Excel.

Deploy ke Railway:
  1. Push folder ini ke GitHub
  2. Buat project baru di railway.app → Deploy from GitHub
  3. Set environment variables (lihat .env.example)
"""

import os
import json
import logging
import sqlite3
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs

from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, filters, ContextTypes,
)

load_dotenv()

BOT_TOKEN   = os.getenv("TELEGRAM_BOT_TOKEN", "")
ALLOWED_UID = int(os.getenv("ALLOWED_USER_ID", "0"))
SYNC_SECRET = os.getenv("SYNC_SECRET", "ganti_ini_dengan_string_acak")
PORT        = int(os.getenv("PORT", 8080))          # Railway inject $PORT otomatis
DB_PATH     = os.getenv("DB_PATH", "transactions.db")

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)

# ── Database ──────────────────────────────────────────────────────────────────

def db_connect():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def db_init():
    con = db_connect()
    con.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at  TEXT    NOT NULL,
            date        TEXT    NOT NULL,
            type        TEXT    NOT NULL,
            category    TEXT    NOT NULL,
            amount      REAL    NOT NULL,
            description TEXT    NOT NULL,
            synced      INTEGER NOT NULL DEFAULT 0
        )
    """)
    con.commit()
    con.close()
    log.info("Database ready: %s", DB_PATH)

def db_insert(date, trans_type, category, amount, description):
    con = db_connect()
    con.execute(
        "INSERT INTO transactions (created_at,date,type,category,amount,description) VALUES (?,?,?,?,?,?)",
        (datetime.now().isoformat(), date, trans_type, category, amount, description),
    )
    con.commit()
    con.close()

def db_get_pending():
    con = db_connect()
    rows = con.execute(
        "SELECT id,date,type,category,amount,description FROM transactions WHERE synced=0 ORDER BY id"
    ).fetchall()
    con.close()
    return rows

def db_mark_synced(ids: list[int]):
    if not ids:
        return
    con = db_connect()
    con.execute(f"UPDATE transactions SET synced=1 WHERE id IN ({','.join('?'*len(ids))})", ids)
    con.commit()
    con.close()

def db_count():
    con = db_connect()
    total, pending = con.execute(
        "SELECT COUNT(*), SUM(CASE WHEN synced=0 THEN 1 ELSE 0 END) FROM transactions"
    ).fetchone()
    con.close()
    return total or 0, pending or 0

# ── Referensi data ────────────────────────────────────────────────────────────

MONTH_SHEETS = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December",
]

VALID_TYPES = {
    "income":  "Income",
    "expanse": "Expanse",
    "invest":  "Invest",
    "savings": "Savings",
}

CATEGORIES = {
    "Income":  ["Gaji","Bonus","Hasil Bisnis","Cashback"],
    "Expanse": ["Makan & Minum","Transportasi","Tagihan","Belanja Bulanan",
                "Internet","iCloud+","Netflix","Claude AI"],
    "Invest":  ["Reksa Dana","Saham","Emas","Deposito"],
    "Savings": ["Dana Darurat","Tabungan Liburan","Tabungan Barang"],
}

CATEGORY_ALIAS = {
    "makan":"Makan & Minum","transport":"Transportasi","tagihan":"Tagihan",
    "belanja":"Belanja Bulanan","internet":"Internet","icloud":"iCloud+",
    "netflix":"Netflix","claude":"Claude AI","gaji":"Gaji","bonus":"Bonus",
    "bisnis":"Hasil Bisnis","cashback":"Cashback","reksadana":"Reksa Dana",
    "reksa":"Reksa Dana","saham":"Saham","emas":"Emas","deposito":"Deposito",
    "darurat":"Dana Darurat","dana_darurat":"Dana Darurat",
    "liburan":"Tabungan Liburan","tabungan":"Tabungan Barang",
}

COLOR_EMOJI = {"Income":"🟢","Expanse":"🔴","Invest":"🟡","Savings":"🔵"}

def current_sheet():
    return MONTH_SHEETS[datetime.now().month - 1]

def resolve_category(raw: str, trans_type: str) -> str | None:
    key = raw.lower().replace(" ", "_")
    if key in CATEGORY_ALIAS:
        return CATEGORY_ALIAS[key]
    for cat in CATEGORIES.get(trans_type, []):
        if cat.lower() == raw.lower():
            return cat
    return None

def format_rupiah(v: float) -> str:
    return f"Rp {v:,.0f}".replace(",", ".")

# ── Guard ─────────────────────────────────────────────────────────────────────

def owner_only(func):
    async def wrapper(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if update.effective_user.id != ALLOWED_UID:
            await update.message.reply_text("⛔ Akses ditolak.")
            return ConversationHandler.END
        return await func(update, ctx)
    wrapper.__name__ = func.__name__
    return wrapper

# ── Handlers ──────────────────────────────────────────────────────────────────

@owner_only
async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📒 *Financial Bot*\n\n"
        "*Quick add:*\n"
        "`/add [type] [kategori] [jumlah] [keterangan]`\n\n"
        "*Contoh:*\n"
        "`/add income gaji 5000000 Gaji bulan ini`\n"
        "`/add expanse makan 50000 Makan siang`\n"
        "`/add invest saham 1000000 Beli BBCA`\n"
        "`/add savings darurat 500000 Nabung rutin`\n\n"
        "*Perintah lain:*\n"
        "/status  — transaksi pending & total\n"
        "/tambah  — input step-by-step\n"
        "/batal   — batalkan input\n\n"
        "💡 Transaksi tersimpan di cloud. Otomatis masuk Excel saat PC nyala.",
        parse_mode="Markdown",
    )

@owner_only
async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    total, pending = db_count()
    await update.message.reply_text(
        f"📊 *Status Database*\n\n"
        f"Total transaksi  : {total}\n"
        f"Belum sync Excel : {pending}\n\n"
        f"Transaksi akan masuk Excel otomatis saat PC kamu nyala.",
        parse_mode="Markdown",
    )

@owner_only
async def cmd_add(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    args = ctx.args
    if len(args) < 3:
        await update.message.reply_text(
            "Format: `/add [type] [kategori] [jumlah] [keterangan]`",
            parse_mode="Markdown",
        )
        return

    raw_type = args[0].lower()
    if raw_type not in VALID_TYPES:
        await update.message.reply_text(
            f"Type tidak valid: `{args[0]}`\nPilihan: income · expanse · invest · savings",
            parse_mode="Markdown",
        )
        return

    trans_type = VALID_TYPES[raw_type]
    category   = resolve_category(args[1], trans_type)
    if category is None:
        await update.message.reply_text(
            f"Kategori `{args[1]}` tidak dikenal.\n"
            f"Pilihan: {' · '.join(CATEGORIES[trans_type])}",
            parse_mode="Markdown",
        )
        return

    try:
        amount = float(args[2].replace(",", "").replace(".", ""))
    except ValueError:
        await update.message.reply_text(f"Jumlah tidak valid: `{args[2]}`", parse_mode="Markdown")
        return

    description = " ".join(args[3:]) if len(args) > 3 else "-"
    today       = datetime.now().strftime("%Y-%m-%d")

    db_insert(today, trans_type, category, amount, description)
    _, pending = db_count()

    await update.message.reply_text(
        f"✅ *Transaksi disimpan!*\n\n"
        f"{COLOR_EMOJI[trans_type]} *{trans_type}* — {category}\n"
        f"💵 {format_rupiah(amount)}\n"
        f"📝 {description}\n"
        f"📅 {datetime.now().strftime('%d/%m/%Y')} · Sheet: {current_sheet()}\n\n"
        f"⏳ Akan masuk Excel saat PC nyala ({pending} pending).",
        parse_mode="Markdown",
    )

# ── Conversational /tambah ────────────────────────────────────────────────────

ASK_TYPE, ASK_CAT, ASK_AMOUNT, ASK_DESC = range(4)

@owner_only
async def cmd_tambah(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📝 *Input Transaksi*\n\nKetik type:\n`income · expanse · invest · savings`",
        parse_mode="Markdown",
    )
    return ASK_TYPE

async def ask_type(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().lower()
    if raw not in VALID_TYPES:
        await update.message.reply_text("Ketik: `income · expanse · invest · savings`", parse_mode="Markdown")
        return ASK_TYPE
    ctx.user_data["type"] = VALID_TYPES[raw]
    cats = " · ".join(CATEGORIES[ctx.user_data["type"]])
    await update.message.reply_text(f"Kategori:\n`{cats}`", parse_mode="Markdown")
    return ASK_CAT

async def ask_cat(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    cat = resolve_category(update.message.text.strip(), ctx.user_data["type"])
    if cat is None:
        await update.message.reply_text(
            f"Tidak dikenal. Pilih:\n`{' · '.join(CATEGORIES[ctx.user_data['type']])}`",
            parse_mode="Markdown",
        )
        return ASK_CAT
    ctx.user_data["category"] = cat
    await update.message.reply_text("Jumlah (contoh: `150000`):", parse_mode="Markdown")
    return ASK_AMOUNT

async def ask_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    try:
        ctx.user_data["amount"] = float(
            update.message.text.strip().replace(",", "").replace(".", "")
        )
    except ValueError:
        await update.message.reply_text("Masukkan angka, contoh: `50000`", parse_mode="Markdown")
        return ASK_AMOUNT
    await update.message.reply_text("Keterangan (atau `-` untuk skip):")
    return ASK_DESC

async def ask_desc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    d = ctx.user_data
    d["description"] = update.message.text.strip()
    today = datetime.now().strftime("%Y-%m-%d")
    db_insert(today, d["type"], d["category"], d["amount"], d["description"])
    _, pending = db_count()
    await update.message.reply_text(
        f"✅ *Tersimpan!*\n\n"
        f"{COLOR_EMOJI[d['type']]} *{d['type']}* — {d['category']}\n"
        f"💵 {format_rupiah(d['amount'])}\n"
        f"📝 {d['description']}\n\n"
        f"⏳ Pending: {pending} transaksi menunggu sync ke Excel.",
        parse_mode="Markdown",
    )
    ctx.user_data.clear()
    return ConversationHandler.END

async def cmd_batal(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text("❌ Dibatalkan.")
    return ConversationHandler.END

# ── HTTP API (untuk sync_to_excel.py) ────────────────────────────────────────

class SyncHandler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass  # suppress default HTTP logs

    def send_json(self, code: int, data):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def check_auth(self) -> bool:
        qs = parse_qs(urlparse(self.path).query)
        return qs.get("token", [""])[0] == SYNC_SECRET

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/health":
            self.send_json(200, {"status": "ok"})
            return
        if path == "/pending":
            if not self.check_auth():
                self.send_json(401, {"error": "Unauthorized"})
                return
            rows = db_get_pending()
            self.send_json(200, [
                {"id": r[0], "date": r[1], "type": r[2],
                 "category": r[3], "amount": r[4], "description": r[5]}
                for r in rows
            ])
            return
        self.send_json(404, {"error": "Not found"})

    def do_POST(self):
        if urlparse(self.path).path == "/mark_synced":
            if not self.check_auth():
                self.send_json(401, {"error": "Unauthorized"})
                return
            length = int(self.headers.get("Content-Length", 0))
            body   = json.loads(self.rfile.read(length))
            ids    = body.get("ids", [])
            db_mark_synced(ids)
            self.send_json(200, {"marked": len(ids)})
            return
        self.send_json(404, {"error": "Not found"})

def start_http_server():
    server = HTTPServer(("0.0.0.0", PORT), SyncHandler)
    log.info("HTTP API berjalan di port %d", PORT)
    server.serve_forever()

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    db_init()

    # HTTP server selalu jalan duluan (Railway health check perlu ini)
    t = threading.Thread(target=start_http_server, daemon=True)
    t.start()

    if not BOT_TOKEN or "isi_token" in BOT_TOKEN:
        log.error("TELEGRAM_BOT_TOKEN belum diset. HTTP server tetap jalan.")
        threading.Event().wait()  # block selamanya agar proses tidak exit
        return
    if ALLOWED_UID == 0:
        log.error("ALLOWED_USER_ID belum diset. HTTP server tetap jalan.")
        threading.Event().wait()
        return

    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("tambah", cmd_tambah)],
        states={
            ASK_TYPE:   [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_type)],
            ASK_CAT:    [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_cat)],
            ASK_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_amount)],
            ASK_DESC:   [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_desc)],
        },
        fallbacks=[CommandHandler("batal", cmd_batal)],
    )

    app.add_handler(CommandHandler("start",  cmd_help))
    app.add_handler(CommandHandler("help",   cmd_help))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("add",    cmd_add))
    app.add_handler(conv)

    log.info("Bot aktif — sheet bulan ini: %s", current_sheet())
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
