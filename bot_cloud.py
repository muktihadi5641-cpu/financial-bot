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
            currency    TEXT    NOT NULL DEFAULT 'IDR',
            description TEXT    NOT NULL,
            synced      INTEGER NOT NULL DEFAULT 0
        )
    """)
    # Migration: tambah kolom currency jika DB lama belum punya
    try:
        con.execute("ALTER TABLE transactions ADD COLUMN currency TEXT NOT NULL DEFAULT 'IDR'")
        log.info("DB migration: kolom currency ditambahkan")
    except Exception:
        pass  # kolom sudah ada
    con.commit()
    con.close()
    log.info("Database ready: %s", DB_PATH)

def db_insert(date, trans_type, category, amount, currency, description):
    con = db_connect()
    con.execute(
        "INSERT INTO transactions (created_at,date,type,category,amount,currency,description) VALUES (?,?,?,?,?,?,?)",
        (datetime.now().isoformat(), date, trans_type, category, amount, currency, description),
    )
    con.commit()
    con.close()

def db_get_pending():
    con = db_connect()
    rows = con.execute(
        "SELECT id,date,type,category,amount,currency,description FROM transactions WHERE synced=0 ORDER BY id"
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

def parse_amount(text: str) -> tuple[float, str] | None:
    """Return (amount, currency) atau None. Currency: 'NTD' atau 'IDR'.
    NTD: '100nt', '500ntd', 'NT$1,500', 'nt 200'
    IDR: '15ribu', '50rb', '1.5jt', '300.000', '50000'
    """
    t = text.lower().strip()

    # ── NTD detection (cek duluan sebelum IDR) ────────────────────────────────
    # "100nt", "500ntd", "1,500nt"
    m = re.search(r'(\d[\d,.]*)\s*(?:ntd|nt)\b', t)
    if not m:
        # "NT$100", "NT 100", "nt100"
        m = re.search(r'\bnt\$?\s*(\d[\d,.]*)', t)
    if m:
        raw = m.group(1)
        # pola ribuan Indonesia: 1.500 atau 1.500.000
        if re.match(r'\d{1,3}(?:\.\d{3})+$', raw):
            raw = raw.replace('.', '')
        else:
            raw = raw.replace(',', '')
        try:
            return float(raw), 'NTD'
        except ValueError:
            pass

    # ── IDR detection ─────────────────────────────────────────────────────────
    t = t.replace('rp', '').replace('idr', '').strip()

    # ribu / rb / k  → ×1.000
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:ribu|rbu|rb|k)(?:\b|$)', t)
    if m:
        return float(m.group(1).replace(',', '.')) * 1_000, 'IDR'

    # juta / jt → ×1.000.000
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:juta|jt)(?:\b|$)', t)
    if m:
        return float(m.group(1).replace(',', '.')) * 1_000_000, 'IDR'

    # angka dengan titik sebagai pemisah ribuan: 15.000 / 1.500.000
    m = re.search(r'\b(\d{1,3}(?:\.\d{3})+)\b', t)
    if m:
        return float(m.group(1).replace('.', '')), 'IDR'

    # angka polos ≥ 4 digit
    m = re.search(r'\b(\d{4,})\b', t)
    if m:
        return float(m.group(1)), 'IDR'

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
    'Beasiswa': ['beasiswa', 'scholarship', 'stipend', 'beasiswaku'],
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

MONTHS_ID: dict[str, int] = {
    'januari':1,'februari':2,'maret':3,'april':4,'mei':5,'juni':6,
    'juli':7,'agustus':8,'september':9,'oktober':10,'november':11,'desember':12,
    'jan':1,'feb':2,'mar':3,'apr':4,'jun':6,'jul':7,
    'agu':8,'ags':8,'sep':9,'okt':10,'nov':11,'des':12,
}

DAYS_ID: dict[str, int] = {
    'senin':0,'selasa':1,'rabu':2,'kamis':3,'jumat':4,'sabtu':5,'minggu':6,
}

_WORD_NUMS = {
    'satu':1,'dua':2,'tiga':3,'empat':4,'lima':5,
    'enam':6,'tujuh':7,'delapan':8,'sembilan':9,'sepuluh':10,
}


def _match(text: str, mapping: dict[str, list[str]]) -> str | None:
    for cat, keywords in mapping.items():
        for kw in keywords:
            if kw in text:
                return cat
    return None


_TYPE_KEYWORDS = {
    'Income':  r'\b(income|pemasukan|pendapatan)\b',
    'Savings': r'\b(saving|savings|saving|tabung|menabung|nabung|simpan)\b',
    'Invest':  r'\b(invest|investasi|investasikan)\b',
    'Expanse': r'\b(expanse|expense|pengeluaran)\b',
}

_TYPE_DEFAULTS = {
    'Income':  ('Income',  'Gaji'),
    'Savings': ('Savings', 'Tabungan Barang'),
    'Invest':  ('Invest',  'Reksa Dana'),
    'Expanse': ('Expanse', 'Belanja Bulanan'),
}

def detect_type_category(text: str) -> tuple[str | None, str | None]:
    t = text.lower()

    # 1. Kata tipe eksplisit (prioritas tertinggi): "income 5jt", "savings 500rb"
    for tipe, pattern in _TYPE_KEYWORDS.items():
        if re.search(pattern, t):
            _maps = {'Income': _INCOME, 'Savings': _SAVINGS,
                     'Invest': _INVEST, 'Expanse': _EXPANSE}
            cat = _match(t, _maps[tipe]) or _TYPE_DEFAULTS[tipe][1]
            return tipe, cat

    # 2. Deteksi dari kata kategori
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


def parse_date(text: str, now: datetime) -> str | None:
    """
    Ekstrak tanggal dari teks Indonesia.
    Return 'YYYY-MM-DD' atau None (pakai tanggal saat pesan dikirim).
    """
    t = text.lower()
    today = now.date()

    # "kemarin lusa" = 2 hari lalu (harus dicek sebelum "kemarin")
    if re.search(r'kemari[en]\s+lusa', t):
        return (today - timedelta(days=2)).strftime('%Y-%m-%d')

    # "kemarin" / "kemaren"
    if re.search(r'\bkemari[en]\b', t):
        return (today - timedelta(days=1)).strftime('%Y-%m-%d')

    # "N hari lalu" / "N hari yang lalu"
    m = re.search(
        r'(\d+|satu|dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|sepuluh)'
        r'\s+hari\s+(?:yang\s+)?lalu', t
    )
    if m:
        raw = m.group(1)
        n = _WORD_NUMS.get(raw) or int(raw)
        return (today - timedelta(days=n)).strftime('%Y-%m-%d')

    # "minggu lalu" = 7 hari lalu (dicek sebelum nama hari "minggu")
    if re.search(r'minggu\s+lalu', t):
        return (today - timedelta(days=7)).strftime('%Y-%m-%d')

    # "tanggal 24" / "tgl 24" / "tanggal 24 april"
    _month_pat = '|'.join(MONTHS_ID.keys())
    m = re.search(
        rf'(?:tanggal|tgl)\s+(\d{{1,2}})(?:\s+({_month_pat}))?', t
    )
    if m:
        day = int(m.group(1))
        if m.group(2):
            month = MONTHS_ID[m.group(2)]
            year = today.year
            if datetime(year, month, day).date() > today:
                year -= 1
        else:
            month = today.month
            year = today.year
            if day > today.day:  # tanggal belum lewat bulan ini → bulan lalu
                month -= 1
                if month == 0:
                    month, year = 12, year - 1
        try:
            return datetime(year, month, day).strftime('%Y-%m-%d')
        except ValueError:
            pass

    # "24 april" / "24 mei" dll.
    m = re.search(rf'\b(\d{{1,2}})\s+({_month_pat})\b', t)
    if m:
        day, month = int(m.group(1)), MONTHS_ID[m.group(2)]
        year = today.year
        try:
            if datetime(year, month, day).date() > today:
                year -= 1
            return datetime(year, month, day).strftime('%Y-%m-%d')
        except ValueError:
            pass

    # Nama hari → occurrence terakhir (termasuk hari ini)
    for day_name, weekday in DAYS_ID.items():
        if re.search(rf'\b{day_name}\b', t):
            days_back = (today.weekday() - weekday) % 7
            return (today - timedelta(days=days_back)).strftime('%Y-%m-%d')

    return None  # tidak ada keterangan waktu → pakai tanggal pesan


def parse_transaction_nl(text: str, msg_date: str, now: datetime) -> dict | None:
    result = parse_amount(text)
    if result is None:
        return None
    amount, currency = result
    trans_type, category = detect_type_category(text)
    if trans_type is None:
        trans_type, category = 'Expanse', 'Belanja Bulanan'
    if category is None:
        category = list(_EXPANSE.keys())[0]
    date = parse_date(text, now) or msg_date
    return {
        'type': trans_type,
        'category': category,
        'amount': amount,
        'currency': currency,
        'description': text.strip(),
        'date': date,
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
    "Income":  ["Gaji","Bonus","Hasil Bisnis","Cashback","Beasiswa"],
    "Expanse": ["Makan & Minum","Transportasi","Tagihan","Belanja Bulanan",
                "Internet","iCloud+","Netflix","Claude AI"],
    "Invest":  ["Reksa Dana","Saham","Emas","Deposito"],
    "Savings": ["Dana Darurat","Tabungan Liburan","Tabungan Barang"],
}

CATEGORY_ALIAS = {
    "makan":"Makan & Minum","transport":"Transportasi","tagihan":"Tagihan",
    "belanja":"Belanja Bulanan","internet":"Internet","icloud":"iCloud+",
    "netflix":"Netflix","claude":"Claude AI","gaji":"Gaji","bonus":"Bonus",
    "bisnis":"Hasil Bisnis","cashback":"Cashback","beasiswa":"Beasiswa","scholarship":"Beasiswa","reksadana":"Reksa Dana",
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

def format_ntd(v: float) -> str:
    return f"NT${v:,.0f}"

def format_amount(v: float, currency: str) -> str:
    return format_ntd(v) if currency == 'NTD' else format_rupiah(v)

# ── Guard ─────────────────────────────────────────────────────────────────────

def owner_only(func):
    async def wrapper(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id if update.effective_user else None
        if uid != ALLOWED_UID:
            return ConversationHandler.END  # silent drop — no response
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
        "*Transaksi NTD (New Taiwan Dollar):*\n"
        "`beli boba 45nt`\n"
        "`bayar sewa 8000ntd`\n"
        "`makan siang 120nt`\n"
        "`NT$1,500 belanja supermarket`\n\n"
        "*Perintah:*\n"
        "/batal — hapus transaksi terakhir\n"
        "/status — ringkasan transaksi\n"
        "/tambah — input step-by-step\n"
        "/help — tampilkan bantuan ini\n\n"
        "💡 Tanggal otomatis sesuai hari kamu kirim pesan.\n"
        "💡 NTD & IDR tersimpan terpisah di Excel.",
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
    now = update.message.date.astimezone(WIB)
    msg_date = now.strftime("%Y-%m-%d")

    parsed = parse_transaction_nl(text, msg_date, now)
    if parsed is None:
        await update.message.reply_text(
            "🤔 Tidak bisa membaca transaksi.\n\n"
            "Pastikan ada nominal, contoh:\n"
            "• `pagi ini beli sarapan 15ribu`\n"
            "• `kemarin bayar bensin 50rb`\n"
            "• `tanggal 24 beli makan 30rb`\n"
            "• `gajian 5jt`\n\n"
            "Atau gunakan /tambah untuk input terpandu.",
            parse_mode="Markdown",
        )
        return

    currency = parsed.get('currency', 'IDR')
    db_insert(parsed['date'], parsed['type'], parsed['category'],
              parsed['amount'], currency, parsed['description'])
    _, pending = db_count()

    date_display = datetime.strptime(parsed['date'], "%Y-%m-%d").strftime("%d/%m/%Y")
    date_note = " _(disesuaikan)_" if parsed['date'] != msg_date else ""
    await update.message.reply_text(
        f"✅ *Tersimpan!*\n\n"
        f"{COLOR_EMOJI[parsed['type']]} *{parsed['type']}* — {parsed['category']}\n"
        f"💵 {format_amount(parsed['amount'], currency)}\n"
        f"📅 {date_display}{date_note}\n\n"
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

@owner_only
async def ask_type(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().lower()
    if raw not in VALID_TYPES:
        await update.message.reply_text("Ketik: `income · expanse · invest · savings`", parse_mode="Markdown")
        return ASK_TYPE
    ctx.user_data["type"] = VALID_TYPES[raw]
    cats = " · ".join(CATEGORIES[ctx.user_data["type"]])
    await update.message.reply_text(f"Kategori:\n`{cats}`", parse_mode="Markdown")
    return ASK_CAT

@owner_only
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

@owner_only
async def ask_amount(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    result = parse_amount(update.message.text.strip())
    if result is None:
        await update.message.reply_text(
            "Masukkan angka, contoh: `50000`, `50rb`, atau `200nt` untuk NTD",
            parse_mode="Markdown",
        )
        return ASK_AMOUNT
    amount, currency = result
    ctx.user_data["amount"] = amount
    ctx.user_data["currency"] = currency
    await update.message.reply_text("Keterangan (atau `-` untuk skip):")
    return ASK_DESC

@owner_only
async def ask_desc(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    d = ctx.user_data
    d["description"] = update.message.text.strip()
    today = datetime.now(WIB).strftime("%Y-%m-%d")
    currency = d.get("currency", "IDR")
    db_insert(today, d["type"], d["category"], d["amount"], currency, d["description"])
    _, pending = db_count()
    await update.message.reply_text(
        f"✅ *Tersimpan!*\n\n"
        f"{COLOR_EMOJI[d['type']]} *{d['type']}* — {d['category']}\n"
        f"💵 {format_amount(d['amount'], currency)}\n\n"
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
                 "category": r[3], "amount": r[4], "currency": r[5], "description": r[6]}
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
