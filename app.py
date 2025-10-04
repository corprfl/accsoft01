import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from datetime import datetime

st.set_page_config(page_title="Aplikasi Laporan Keuangan Profesional", layout="wide")

# =========================
# Fungsi bantu
# =========================
def rename_cols(df):
    df.columns = df.columns.astype(str).str.strip().str.lower()
    mapping = {
        "kode akun": "kode_akun",
        "kode": "kode_akun",
        "akun": "kode_akun",
        "saldo awal": "saldo",
        "saldo_awal": "saldo",
        "nilai": "saldo",
        "saldo akhir": "saldo_akhir",
        "posisi normal akun": "posisi_normal_akun",
        "posisi normal": "posisi_normal_akun",
        "posisi": "posisi_normal_akun"
    }
    df = df.rename(columns={k: v for k, v in mapping.items() if k in df.columns})
    return df

def hitung_saldo(saldo_awal, debit, kredit, posisi):
    if str(posisi).lower() == "debit":
        return saldo_awal + debit - kredit
    else:
        return saldo_awal - debit + kredit

def format_rp(x):
    return f"Rp {x:,.0f}".replace(",", ".")

# =========================
# PDF LABA RUGI
# =========================
def export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    margin_x = 2*cm
    y = h - 3*cm
    line_height = 14

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, h-2*cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, h-2.7*cm, "LAPORAN LABA RUGI")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, h-3.3*cm, f"Untuk Periode yang Berakhir Pada {periode_text}")

    top_y = h-3.6*cm
    bottom_y = 3*cm
    c.rect(margin_x-0.5*cm, bottom_y-0.5*cm, w-2*(margin_x-0.5*cm), top_y-bottom_y+0.5*cm)

    def tulis_baris(label, amount=None, bold=False, gap=0):
        nonlocal y
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        c.drawString(margin_x, y, str(label))
        if amount is not None:
            c.drawRightString(w-margin_x, y, f"{format_rp(amount)}")
        y -= (line_height + gap)

    def tulis_total(label, amount):
        nonlocal y
        c.line(margin_x, y+3, w-margin_x, y+3)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margin_x, y, str(label))
        c.drawRightString(w-margin_x, y, f"{format_rp(amount)}")
        y -= line_height
        c.line(margin_x, y+line_height-3, w-margin_x, y+line_height-3)

    sections = df_laba["sub_tipe_laporan"].unique()
    for section in sections:
        subset = df_laba[df_laba["sub_tipe_laporan"] == section]
        tulis_baris(section, bold=True)
        for _, r in subset.iterrows():
            tulis_baris("   " + str(r["nama_akun"]), r["saldo_akhir_adj"])
        tulis_total(f"TOTAL {section.upper()}", subset["saldo_akhir_adj"].sum())
        y -= 10

    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin_x, y, "LABA (RUGI) BERSIH")
    c.drawRightString(w-margin_x, y, f"{format_rp(laba_bersih)}")
    c.line(w-margin_x-180, y-3, w-margin_x, y-3)
    c.line(w-margin_x-180, y-6, w-margin_x, y-6)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf

# =========================
# PDF NERACA
# =========================
def export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    margin_x = 2*cm
    y = h - 3*cm
    line_height = 14

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, h-2*cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, h-2.7*cm, "LAPORAN POSISI KEUANGAN")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, h-3.3*cm, f"Per {periode_text}")

    top_y = h-3.6*cm
    bottom_y = 3*cm
    c.rect(margin_x-0.5*cm, bottom_y-0.5*cm, w-2*(margin_x-0.5*cm), top_y-bottom_y+0.5*cm)

    def tulis_baris(label, amount=None, bold=False, gap=0):
        nonlocal y
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        c.drawString(margin_x, y, str(label))
        if amount is not None:
            c.drawRightString(w-margin_x, y, f"{format_rp(amount)}")
        y -= (line_height + gap)

    def tulis_total(label, amount):
        nonlocal y
        c.line(margin_x, y+3, w-margin_x, y+3)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margin_x, y, str(label))
        c.drawRightString(w-margin_x, y, f"{format_rp(amount)}")
        y -= line_height
        c.line(margin_x, y+line_height-3, w-margin_x, y+line_height-3)

    tulis_baris("ASET", bold=True)
    for _, r in df_aset.iterrows():
        tulis_baris("   " + str(r["nama_akun"]), r["saldo_akhir_adj"])
    tulis_total("TOTAL ASET", total_aset)
    y -= 10

    tulis_baris("KEWAJIBAN", bold=True)
    for _, r in df_kewajiban.iterrows():
        tulis_baris("   " + str(r["nama_akun"]), r["saldo_akhir_adj"])
    tulis_total("TOTAL KEWAJIBAN", total_kewajiban)
    y -= 10

    tulis_baris("EKUITAS", bold=True)
    for _, r in df_ekuitas.iterrows():
        tulis_baris("   " + str(r["nama_akun"]), r["saldo_akhir_adj"])
    tulis_total("TOTAL EKUITAS", total_ekuitas)
    y -= 10

    c.setFont("Helvetica-Bold", 10)
    c.drawString(margin_x, y, "TOTAL KEWAJIBAN + EKUITAS")
    c.drawRightString(w-margin_x, y, f"{format_rp(total_kewajiban + total_ekuitas)}")
    c.line(w-margin_x-180, y-3, w-margin_x, y-3)
    c.line(w-margin_x-180, y-6, w-margin_x, y-6)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf

# =========================
# STREAMLIT
# =========================
st.title("📘 Generator Laporan Keuangan Profesional")

coa_file = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])

nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
tanggal_awal = st.date_input("Tanggal Awal Periode", datetime(2025,1,1))
tanggal_akhir = st.date_input("Tanggal Akhir Periode", datetime(2025,12,31))
periode_text = tanggal_akhir.strftime("%d %B %Y")

if not (coa_file and saldo_file and jurnal_file):
    st.warning("⚠️ Silakan upload ketiga file terlebih dahulu.")
    st.stop()

coa = rename_cols(pd.read_excel(coa_file))
saldo_awal = rename_cols(pd.read_excel(saldo_file))
jurnal = rename_cols(pd.read_excel(jurnal_file))

for df, name in [(coa, "COA"), (saldo_awal, "Saldo Awal"), (jurnal, "Jurnal")]:
    if "kode_akun" not in df.columns:
        st.error(f"❌ File {name} tidak memiliki kolom 'Kode Akun'")
        st.write("Kolom tersedia:", list(df.columns))
        st.stop()

if "posisi_normal_akun" not in coa.columns:
    st.error("❌ Kolom 'Posisi Normal Akun' tidak ditemukan di COA.")
    st.write("Kolom tersedia:", list(coa.columns))
    st.stop()

# === Proses data
jurnal_sum = jurnal.groupby("kode_akun")[["debit","kredit"]].sum().reset_index()
df = coa.merge(saldo_awal[["kode_akun","saldo"]], on="kode_akun", how="left").fillna(0)
df = df.merge(jurnal_sum, on="kode_akun", how="left").fillna(0)

df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(r["saldo"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1)

df["saldo_akhir_adj"] = df.apply(
    lambda r: r["saldo_akhir"] if (
        (r["laporan"].lower() == "aset" and r["posisi_normal_akun"].lower() == "debit")
        or (r["laporan"].lower() in ["kewajiban","ekuitas"] and r["posisi_normal_akun"].lower() == "kredit")
    ) else -r["saldo_akhir"], axis=1
)

df_laba = df[df["laporan"].str.contains("laba", case=False, na=False)]
df_aset = df[df["laporan"].str.contains("aset", case=False, na=False)]
df_kewajiban = df[df["laporan"].str.contains("kewajiban", case=False, na=False)]
df_ekuitas = df[df["laporan"].str.contains("ekuitas", case=False, na=False)]

# === Handle jika kolom sub_tipe_laporan tidak ada
if "sub_tipe_laporan" not in df_laba.columns:
    st.warning("⚠️ Kolom 'sub_tipe_laporan' tidak ditemukan, sistem akan menggunakan kategori otomatis.")
    df_laba["sub_tipe_laporan"] = "Pendapatan"

def get_total(df, keyword):
    return df[df["sub_tipe_laporan"].str.contains(keyword, case=False, na=False)]["saldo_akhir_adj"].sum()

total_pendapatan = get_total(df_laba, "pendapatan")
total_beban_umum = get_total(df_laba, "beban umum")
total_pendapatan_luar = get_total(df_laba, "pendapatan luar")
total_beban_luar = get_total(df_laba, "beban luar")

laba_bersih = total_pendapatan - total_beban_umum + total_pendapatan_luar - total_beban_luar

if "3004" in df_ekuitas["kode_akun"].astype(str).values:
    df_ekuitas.loc[df_ekuitas["kode_akun"].astype(str)=="3004","saldo_akhir_adj"] = laba_bersih

total_aset = df_aset["saldo_akhir_adj"].sum()
total_kewajiban = df_kewajiban["saldo_akhir_adj"].sum()
total_ekuitas = df_ekuitas["saldo_akhir_adj"].sum()

st.header("📈 Laporan Laba Rugi")
st.success(f"💰 Laba (Rugi) Bersih: {format_rp(laba_bersih)}")

st.header("📊 Neraca (Posisi Keuangan)")
st.info(f"Total Aset: {format_rp(total_aset)} | Total Kewajiban + Ekuitas: {format_rp(total_kewajiban + total_ekuitas)}")

# === Export Excel
def export_excel():
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_laba.to_excel(writer, sheet_name="Laba Rugi", index=False)
        df_aset.to_excel(writer, sheet_name="Aset", index=False)
        df_kewajiban.to_excel(writer, sheet_name="Kewajiban", index=False)
        df_ekuitas.to_excel(writer, sheet_name="Ekuitas", index=False)
    buf.seek(0)
    return buf

st.download_button("📥 Download Excel", data=export_excel(), file_name="Laporan_Keuangan.xlsx")

# === Export PDF
pdf_laba = export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text)
st.download_button("📄 Download PDF Laba Rugi", data=pdf_laba, file_name="Laporan_Laba_Rugi.pdf", mime="application/pdf")

pdf_neraca = export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text)
st.download_button("📄 Download PDF Neraca", data=pdf_neraca, file_name="Laporan_Posisi_Keuangan.pdf", mime="application/pdf")
