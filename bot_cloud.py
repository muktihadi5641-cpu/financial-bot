"""
Bot Telegram untuk Financial Management — versi CLOUD (Railway.app)
Berjalan 24/7 di server cloud. Transaksi disimpan ke SQLite.
Saat PC user nyala, sync_to_excel.py mengambil data ini dan menulis ke Excel.
"""

import os
import re
import json
import logging
import sqlite3
import threading
from datetime import datetime, timezone, timedelta
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
PORT        = int(os.getenv("PORT", 8080))
DB_PATH     = os.getenv("DB_PATH", "transactions.db")

WIB = timezone(timedelta(hours=8))

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

def db_delete_last():
    """Hapus transaksi terakhir yang belum disync (untuk /batal)."""
    con = db_connect()
    row = con.execute(
        "SELECT id FROM transactions WHERE synced=0 ORDER BY id DESC LIMIT 1"
    ).fetchone()
    if row:
        con.execute("DELETE FROM transactions WHERE id=?", (row[0],))
        con.commit()
        con.close()
        return True
    con.close()
    return False

# ── NLP Parser ────────────────────────────────────────────────────────────────

def parse_amount(text: str) -> float | None:
    """Parse nominal uang dari teks Indonesia: '15ribu', '50rb', '1.5jt', '300.000', '50000'."""
    t = text.lower().replace('rp', '').replace('idr', '').strip()

    # ribu / rb / k  → ×1.000
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:ribu|rbu|rb|k)(?:\b|$)', t)
    if m:
        return float(m.group(1).replace(',', '.')) * 1_000

    # juta / jt → ×1.000.000
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:juta|jt)(?:\b|$)', t)
    if m:
        return float(m.group(1).replace(',', '.')) * 1_000_000

    # angka dengan titik sebagai pemisah ribuan: 15.000 / 1.500.000
    m = re.search(r'\b(\d{1,3}(?:\.\d{3})+)\b', t)
    if m:
        return float(m.group(1).replace('.', ''))

    # angka polos ≥ 4 digit
    m = re.search(r'\b(\d{4,})\b', t)
    if m:
        return float(m.group(1))

    return None


_EXPANSE: dict[str, list[str]] = {
    'Makan & Minum': [
        'makan', 'sarapan', 'breakfast', 'lunch', 'dinner', 'minum', 'kopi',
        'teh', 'susu', 'nasi', 'ayam', 'bakso', 'soto', 'mie', 'burger',
        'pizza', 'jajan', 'warung', 'resto', 'restoran', 'cafe', 'kantin',
        'camilan', 'snack', 'boba', 'es krim', 'kue', 'roti', 'gorengan',
        'martabak', 'gado', 'lauk', 'nongkrong', 'minum kopi',
    ],
    'Transportasi': [
        'ojek', 'gojek', 'grab', 'bensin', 'bbm', 'parkir', 'busway',
        'kereta', 'mrt', 'lrt', 'bis', 'bus', 'angkot', 'taxi', 'taksi',
        'tol', 'transport', 'krl', 'commuterline', 'naik motor',
        'servis motor', 'ganti oli', 'tambal ban',
    ],
    'Tagihan': [
        'listrik', 'pln', 'pdam', 'air', 'telpon', 'telepon', 'cicilan',
        'tagihan', 'angsuran', 'kredit', 'pulsa',
    ],
    'Belanja Bulanan': [
        'supermarket', 'indomaret', 'alfamart', 'hypermart', 'lottemart',
        'grocery', 'belanja', 'deterjen', 'sabun', 'shampoo', 'pasta gigi',
        'sembako', 'kebutuhan',
    ],
    'Internet': ['wifi', 'internet', 'kuota', 'data', 'indihome', 'firstmedia'],
    'iCloud+': ['icloud'],
    'Netflix': ['netflix'],
    'Claude AI': ['claude', 'anthropic'],
}

_INCOME: dict[str, list[str]] = {
    'Gaji': ['gaji', 'salary', 'upah', 'honor', 'gajian'],
    'Bonus': ['bonus', 'thr', 'insentif', 'komisi'],
    'Hasil Bisnis': ['bisnis', 'usaha', 'jualan', 'dagangan', 'omzet', 'dividen'],
    'Cashback': ['cashback', 'refund', 'pengembalian'],
}

_INVEST: dict[str, list[str]] = {
    'Reksa Dana': ['reksa dana', 'reksadana'],
    'Saham': ['saham', 'stock'],
    'Emas': ['emas', 'logam mulia', 'antam'],
    'Deposito': ['deposito'],
}

_SAVINGS: dict[str, list[str]] = {
    'Dana Darurat': ['dana darurat', 'darurat', 'emergency'],
    'Tabungan Liburan': ['tabungan liburan', 'liburan'],
    'Tabungan Barang': ['nabung', 'menabung', 'tabungan', 'simpan', 'celengan'],
}

_EXPANSE_SIGNALS = ['beli', 'bayar', 'habis', 'keluar', 'traktir', 'order', 'pesan', 'checkout']


def _match(text: str, mapping: dict[str, list[str]]) -> str | None:
    for cat, keywords in mapping.items():
        for kw in keywords:
            if kw in text:
                return cat
    return None


def detect_type_category(text: str) -> tuple[str | None, str | None]:
    t = text.lower()
    # Income & invest lebih spesifik → dicek lebih dulu
    cat = _match(t, _INCOME)
    if cat:
        return 'Income', cat
    cat = _match(t, _INVEST)
    if cat:
        return 'Invest', cat
    cat = _match(t, _SAVINGS)
    if cat:
        return 'Savings', cat
    cat = _match(t, _EXPANSE)
    if cat:
        return 'Expanse', cat
    if any(sig in t for sig in _EXPANSE_SIGNALS):
        return 'Expanse', 'Belanja Bulanan'
    return None, None


def parse_transaction_nl(text: str, msg_date: str) -> dict | None:
    amount = parse_amount(text)
    if amount is None:
        return None
    trans_type, category = detect_type_category(text)
    if trans_type is None:
        # Ada nominal tapi tipe tidak dikenal → default Expanse
        trans_type, category = 'Expanse', 'Belanja Bulanan'
    if category is None:
        category = list(_EXPANSE.keys())[0]
    return {
        'type': trans_type,
        'category': category,
        'amount': amount,
        'description': text.strip(),
        'date': msg_date,
    }

# ── Referensi data ────────────────────────────────────────────────────────────

MONTH_SHEETS = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December",
]

VALID_TYPES = {
    "income": "Income", "expanse": "Expanse",
    "invest": "Invest", "savings": "Savings",
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
    return MONTH_SHEETS[datetime.now(WIB).month - 1]

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
async def cmd_help(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📒 *Financial Bot*\n\n"
        "Cukup ketik transaksimu secara natural:\n"
        "`pagi ini beli sarapan 15ribu`\n"
        "`bayar bensin 50rb`\n"
        "`gajian bulan ini 5jt`\n"
        "`nabung dana darurat 500000`\n"
        "`beli saham BBCA 1.5jt`\n"
        "`bayar listrik 300.000`\n\n"
        "*Perintah:*\n"
        "/batal — hapus transaksi terakhir\n"
        "/status — ringkasan transaksi\n"
        "/tambah — input step-by-step\n"
        "/help — tampilkan bantuan ini\n\n"
        "💡 Tanggal otomatis sesuai hari kamu kirim pesan.",
        parse_mode="Markdown",
    )

@owner_only
async def cmd_status(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    total, pending = db_count()
    await update.message.reply_text(
        f"📊 *Status Database*\n\n"
        f"Total transaksi  : {total}\n"
        f"Belum sync Excel : {pending}\n\n"
        f"Transaksi masuk Excel otomatis saat PC nyala.",
        parse_mode="Markdown",
    )

@owner_only
async def handle_free_text(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    # Tanggal = waktu pesan dikirim, dikonversi ke WIB
    msg_date = update.message.date.astimezone(WIB).strftime("%Y-%m-%d")

    parsed = parse_transaction_nl(text, msg_date)
    if parsed is None:
        await update.message.reply_text(
            "🤔 Tidak bisa membaca transaksi.\n\n"
            "Pastikan ada nominal, contoh:\n"
            "• `pagi ini beli sarapan 15ribu`\n"
            "• `bayar bensin 50rb`\n"
            "• `gajian 5jt`\n\n"
            "Atau gunakan /tambah untuk input terpandu.",
            parse_mode="Markdown",
        )
        return

    db_insert(parsed['date'], parsed['type'], parsed['category'],
              parsed['amount'], parsed['description'])
    _, pending = db_count()

    date_display = datetime.strptime(parsed['date'], "%Y-%m-%d").strftime("%d/%m/%Y")
    await update.message.reply_text(
        f"✅ *Tersimpan!*\n\n"
        f"{COLOR_EMOJI[parsed['type']]} *{parsed['type']}* — {parsed['category']}\n"
        f"💵 {format_rupiah(parsed['amount'])}\n"
        f"📅 {date_display}\n\n"
        f"⏳ {pending} transaksi menunggu sync ke Excel.\n"
        f"_Salah? Ketik /batal_",
        parse_mode="Markdown",
    )

@owner_only
async def cmd_batal(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    if db_delete_last():
        await update.message.reply_text("↩️ Transaksi terakhir dihapus.")
    else:
        await update.message.reply_text("Tidak ada transaksi yang bisa dibatalkan.")
    return ConversationHandler.END

# ── Conversational /tambah ────────────────────────────────────────────────────

ASK_TYPE, ASK_CAT, ASK_AMOUNT, ASK_DESC = range(4)

@owner_only
async def cmd_tambah(update: Update, _ctx: ContextTypes.DEFAULT_TYPE):
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
    await update.message.reply_text("Jumlah (contoh: `150000` atau `150rb`):", parse_mode="Markdown")
    return ASK_AMOUNT

async def ask_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    amount = parse_amount(update.message.text.strip())
    if amount is None:
        await update.message.reply_text("Masukkan angka, contoh: `50000` atau `50rb`", parse_mode="Markdown")
        return ASK_AMOUNT
    ctx.user_data["amount"] = amount
    await update.message.reply_text("Keterangan (atau `-` untuk skip):")
    return ASK_DESC

async def ask_desc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    d = ctx.user_data
    d["description"] = update.message.text.strip()
    today = datetime.now(WIB).strftime("%Y-%m-%d")
    db_insert(today, d["type"], d["category"], d["amount"], d["description"])
    _, pending = db_count()
    await update.message.reply_text(
        f"✅ *Tersimpan!*\n\n"
        f"{COLOR_EMOJI[d['type']]} *{d['type']}* — {d['category']}\n"
        f"💵 {format_rupiah(d['amount'])}\n\n"
        f"⏳ {pending} transaksi menunggu sync ke Excel.",
        parse_mode="Markdown",
    )
    ctx.user_data.clear()
    return ConversationHandler.END

# ── HTTP API (untuk sync_to_excel.py) ────────────────────────────────────────

class SyncHandler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

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

    # HTTP server selalu jalan duluan (Railway health check)
    t = threading.Thread(target=start_http_server, daemon=True)
    t.start()

    if not BOT_TOKEN or "isi_token" in BOT_TOKEN:
        log.error("TELEGRAM_BOT_TOKEN belum diset. HTTP server tetap jalan.")
        threading.Event().wait()
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

    app.add_handler(conv)
    app.add_handler(CommandHandler("start",  cmd_help))
    app.add_handler(CommandHandler("help",   cmd_help))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("batal",  cmd_batal))
    # Free-text handler: semua pesan non-command diparse sebagai transaksi
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_free_text))

    log.info("Bot aktif — sheet bulan ini: %s", current_sheet())
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
