"""
fix_excel.py — perbaiki Financial Management.xlsx:
  1. Conditional formatting: warna opaque + range sesuai tabel (B:T)
  2. SUMIF formula di setiap sheet bulan (diperbaiki dari range salah)
  3. Main Dashboard: link akumulasi semua bulan
  4. Re-apply data validations (openpyxl drop validasi cross-sheet saat save)
"""

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

EXCEL_PATH = r"C:\Claude\Project\Financial\Financial Management.xlsx"

MONTHS = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December",
]
MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]

# ── Warna baris per tipe (ARGB: FF = opaque) ──────────────────────────────────
TYPE_FILLS = {
    "Income":  PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid"),
    "Expanse": PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid"),
    "Invest":  PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid"),
    "Savings": PatternFill(start_color="FFBDD7EE", end_color="FFBDD7EE", fill_type="solid"),
}

# Range sesuai tabel: B(date) C:E(type) F:H(cat) I:J(amount) K:T(desc)
CF_RANGE = "B10:T1000"

# Kategori semua tipe (untuk dropdown)
ALL_CATEGORIES = (
    "Gaji,Bonus,Hasil Bisnis,Cashback,"
    "Makan & Minum,Transportasi,Tagihan,Belanja Bulanan,"
    "Internet,iCloud+,Netflix,Claude AI,"
    "Reksa Dana,Saham,Emas,Deposito,"
    "Dana Darurat,Tabungan Liburan,Tabungan Barang"
)

print("Membuka file Excel...")
wb = openpyxl.load_workbook(EXCEL_PATH)

# ── Fix setiap sheet bulan ─────────────────────────────────────────────────────
for month in MONTHS:
    if month not in wb.sheetnames:
        print(f"  [skip] sheet '{month}' tidak ditemukan")
        continue

    ws = wb[month]
    print(f"  [{month}] fix CF + SUMIF + validasi...")

    # 1. Hapus CF lama, apply ulang dengan warna benar dan range yang tepat
    ws.conditional_formatting = openpyxl.formatting.formatting.ConditionalFormattingList()
    for type_name, fill in TYPE_FILLS.items():
        ws.conditional_formatting.add(
            CF_RANGE,
            FormulaRule(formula=[f'$C10="{type_name}"'], fill=fill)
        )

    # 2. SUMIF IDR (kolom I) — row 4
    ws["E4"] = '=SUMIF($C$10:$C$1000,"Income",$I$10:$I$1000)'
    ws["I4"] = '=SUMIF($C$10:$C$1000,"Expanse",$I$10:$I$1000)'
    ws["M4"] = '=SUMIF($C$10:$C$1000,"Invest",$I$10:$I$1000)'
    ws["Q4"] = '=SUMIF($C$10:$C$1000,"Savings",$I$10:$I$1000)'

    # SUMIF NTD (kolom U = col 21) — row 6 (E4:H5 adalah satu merged cell, row 6 bebas)
    ws["E6"] = '=SUMIF($C$10:$C$1000,"Income",$U$10:$U$1000)'
    ws["I6"] = '=SUMIF($C$10:$C$1000,"Expanse",$U$10:$U$1000)'
    ws["M6"] = '=SUMIF($C$10:$C$1000,"Invest",$U$10:$U$1000)'
    ws["Q6"] = '=SUMIF($C$10:$C$1000,"Savings",$U$10:$U$1000)'

    # Migrasi data NTD yang sudah ada: update formula display di col I
    # agar baris NTD tampil "NT$ X" bukan "Rp -"
    for r in range(10, 1001):
        u_val = ws.cell(row=r, column=21).value
        if u_val and not isinstance(u_val, str) and float(u_val) > 0:
            ws.cell(row=r, column=9).value = f'=IF(U{r}>0,TEXT(U{r},"NT$ #,##0"),$E$8)'

    # Label kolom backing NTD di row 9, col U (backing column untuk SUMIF)
    ws.cell(row=9, column=21).value = "NTD_DATA"

    # 3. Re-apply data validations (openpyxl drop cross-sheet ref saat load)
    ws.data_validations.dataValidation = []

    # TYPE TRANSACTION dropdown — col C
    dv_type = DataValidation(
        type="list",
        formula1="Refference!$B$2:$E$2",
        allow_blank=True,
        showDropDown=False,
    )
    dv_type.sqref = "C10:C1000"
    ws.add_data_validation(dv_type)

    # CATEGORY dropdown — col F (kompatibel tanpa cross-sheet dependency)
    dv_cat = DataValidation(
        type="list",
        formula1=f'"{ALL_CATEGORIES}"',
        allow_blank=True,
        showDropDown=False,
    )
    dv_cat.sqref = "F10:F1000"
    ws.add_data_validation(dv_cat)

# ── Fix Main Dashboard — akumulasi semua bulan ────────────────────────────────
print("  [Main Dashboard] link semua bulan...")
ws_main = wb["Main Dashboard"]

for i, month in enumerate(MONTHS):
    row = 9 + i  # Jan=9, Feb=10, ..., Dec=20
    # IDR (D:E merged → anchor D, col 4)
    ws_main.cell(row=row, column=4).value  = f"={month}!E4"   # D = Income IDR
    ws_main.cell(row=row, column=8).value  = f"={month}!I4"   # H = Expanse IDR
    ws_main.cell(row=row, column=12).value = f"={month}!M4"   # L = Invest IDR
    ws_main.cell(row=row, column=16).value = f"={month}!Q4"   # P = Savings IDR
    # NTD — kolom S,T,U,V (19,20,21,22) — bebas dari semua merge yang ada
    ws_main.cell(row=row, column=19).value = f"={month}!E6"   # S = Income NTD
    ws_main.cell(row=row, column=20).value = f"={month}!I6"   # T = Expanse NTD
    ws_main.cell(row=row, column=21).value = f"={month}!M6"   # U = Invest NTD
    ws_main.cell(row=row, column=22).value = f"={month}!Q6"   # V = Savings NTD

# Total row 21 — IDR
ws_main["D21"] = "=SUM(D9:D20)"
ws_main["H21"] = "=SUM(H9:H20)"
ws_main["L21"] = "=SUM(L9:L20)"
ws_main["P21"] = "=SUM(P9:P20)"

# Total row 21 — NTD
ws_main.cell(row=21, column=19).value = "=SUM(S9:S20)"   # S = Income NTD total
ws_main.cell(row=21, column=20).value = "=SUM(T9:T20)"   # T = Expanse NTD total
ws_main.cell(row=21, column=21).value = "=SUM(U9:U20)"   # U = Invest NTD total
ws_main.cell(row=21, column=22).value = "=SUM(V9:V20)"   # V = Savings NTD total

# Summary cards IDR row 5 (sudah ada, tulis ulang untuk safety)
ws_main["C5"] = "=D21"
ws_main["G5"] = "=H21"
ws_main["K5"] = "=L21"
ws_main["O5"] = "=P21"

# Header NTD di row 7 dan 8 (rows 5-6 = IDR summary merges, row 7-8 bebas di col S+)
ws_main.cell(row=7, column=19).value = "NTD SUMMARY"
ws_main.cell(row=8, column=19).value = "Income"
ws_main.cell(row=8, column=20).value = "Expanse"
ws_main.cell(row=8, column=21).value = "Invest"
ws_main.cell(row=8, column=22).value = "Savings"

# Summary cards NTD di row 5-6 bebas di col S (19+)
ws_main.cell(row=5, column=19).value = "=S21"   # Income NTD total
ws_main.cell(row=5, column=20).value = "=T21"   # Expanse NTD total
ws_main.cell(row=5, column=21).value = "=U21"   # Invest NTD total
ws_main.cell(row=5, column=22).value = "=V21"   # Savings NTD total

# ── Save ──────────────────────────────────────────────────────────────────────
print("Menyimpan...")
try:
    wb.save(EXCEL_PATH)
    print("Selesai! File tersimpan.")
except PermissionError:
    print("ERROR: File Excel sedang terbuka. Tutup Excel dulu lalu jalankan ulang.")
