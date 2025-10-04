import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors

# Judul aplikasi
st.title("📊 Aplikasi Laporan Keuangan")

# Upload file
coa_file = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])

if not (coa_file and saldo_file and jurnal_file):
    st.warning("⚠️ Silakan upload ketiga file: COA, Saldo Awal, dan Jurnal")
    st.stop()

# Baca file
coa = pd.read_excel(coa_file)
saldo_awal = pd.read_excel(saldo_file)
jurnal = pd.read_excel(jurnal_file)

# Normalisasi kolom
coa.columns = coa.columns.str.strip().str.lower()
saldo_awal.columns = saldo_awal.columns.str.strip().str.lower()
jurnal.columns = jurnal.columns.str.strip().str.lower()

# Pastikan numeric
if "saldo" in saldo_awal.columns:
    saldo_awal["saldo"] = pd.to_numeric(saldo_awal["saldo"], errors="coerce").fillna(0)

jurnal["debit"] = pd.to_numeric(jurnal["debit"], errors="coerce").fillna(0)
jurnal["kredit"] = pd.to_numeric(jurnal["kredit"], errors="coerce").fillna(0)

# Gabungkan data
df = coa.merge(saldo_awal, on="kode_akun", how="left").fillna(0)

# Hitung saldo akhir
def hitung_saldo(saldo_awal, debit, kredit, posisi):
    if posisi.lower() == "debit":
        return saldo_awal + debit - kredit
    else:
        return saldo_awal - debit + kredit

# Total debit/kredit dari jurnal
total_jurnal = jurnal.groupby("kode_akun")[["debit", "kredit"]].sum().reset_index()
df = df.merge(total_jurnal, on="kode_akun", how="left").fillna(0)

df["saldo_akhir"] = df.apply(
    lambda r: hitung_saldo(r["saldo"], r["debit"], r["kredit"], r["posisi_normal_akun"]),
    axis=1,
)

# Pisahkan laporan
laba_rugi = df[df["laporan"].str.contains("Laba Rugi", case=False, na=False)]
neraca = df[df["laporan"].str.contains("Posisi Keuangan", case=False, na=False)]

# === LABA RUGI ===
pendapatan = laba_rugi[laba_rugi["sub_tipe_laporan"].str.contains("Pendapatan", case=False, na=False)]
beban_umum = laba_rugi[laba_rugi["sub_tipe_laporan"].str.contains("Beban Umum", case=False, na=False)]
pendapatan_luar = laba_rugi[laba_rugi["sub_tipe_laporan"].str.contains("Pendapatan Luar", case=False, na=False)]
beban_luar = laba_rugi[laba_rugi["sub_tipe_laporan"].str.contains("Beban Luar", case=False, na=False)]

total_pendapatan = pendapatan["saldo_akhir"].sum()
total_beban_umum = beban_umum["saldo_akhir"].sum()
total_pendapatan_luar = pendapatan_luar["saldo_akhir"].sum()
total_beban_luar = beban_luar["saldo_akhir"].sum()

laba_bersih = total_pendapatan - total_beban_umum + total_pendapatan_luar - total_beban_luar

# === NERACA ===
aset = neraca[neraca["sub_tipe_laporan"].str.contains("Aset", case=False, na=False)].copy()
kewajiban = neraca[neraca["sub_tipe_laporan"].str.contains("Kewajiban", case=False, na=False)].copy()
ekuitas = neraca[neraca["sub_tipe_laporan"].str.contains("Ekuitas", case=False, na=False)].copy()

# Rule saldo normal
def adjust_saldo(row):
    if row["posisi_normal_akun"].lower() == "debit":
        return row["saldo_akhir"]
    else:
        return -row["saldo_akhir"]

aset["saldo_akhir_adj"] = aset.apply(adjust_saldo, axis=1)
kewajiban["saldo_akhir_adj"] = kewajiban.apply(adjust_saldo, axis=1)
ekuitas["saldo_akhir_adj"] = ekuitas.apply(adjust_saldo, axis=1)

# Tambahkan laba bersih ke ekuitas (Saldo Laba Berjalan)
ekuitas = pd.concat(
    [ekuitas, pd.DataFrame([{"kode_akun": "3004", "nama_akun": "Saldo Laba (Rugi) Berjalan", "saldo_akhir_adj": laba_bersih}])]
)

total_aset = aset["saldo_akhir_adj"].sum()
total_kewajiban = kewajiban["saldo_akhir_adj"].sum()
total_ekuitas = ekuitas["saldo_akhir_adj"].sum()

# === TAMPILKAN DI STREAMLIT ===
st.subheader("📑 Laporan Laba Rugi")
st.write("**LABA (RUGI) BERSIH : Rp {:,.0f}**".format(laba_bersih))

st.subheader("📑 Laporan Posisi Keuangan (Neraca)")
st.write("**TOTAL ASET : Rp {:,.0f}**".format(total_aset))
st.write("**TOTAL KEWAJIBAN : Rp {:,.0f}**".format(total_kewajiban))
st.write("**TOTAL EKUITAS : Rp {:,.0f}**".format(total_ekuitas))

# === EXPORT EXCEL ===
def export_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        laba_rugi.to_excel(writer, sheet_name="Laba Rugi", index=False)
        aset.to_excel(writer, sheet_name="Aset", index=False)
        kewajiban.to_excel(writer, sheet_name="Kewajiban", index=False)
        ekuitas.to_excel(writer, sheet_name="Ekuitas", index=False)
    return output.getvalue()

st.download_button(
    "📥 Export ke Excel",
    data=export_excel(),
    file_name="laporan_keuangan.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# === EXPORT PDF ===
def export_pdf_laba_rugi():
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 2 * cm

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, "PT Contoh Sejahtera")
    y -= 1*cm
    c.drawCentredString(w/2, y, "LAPORAN LABA RUGI")
    y -= 0.7*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, "Untuk Periode yang Berakhir Pada 31 Desember 2025")
    y -= 1*cm

    # Pendapatan
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "Pendapatan")
    y -= 0.5*cm
    c.setFont("Helvetica", 10)
    c.drawRightString(w-2*cm, y, f"Rp {total_pendapatan:,.0f}")
    y -= 0.5*cm
    c.line(2*cm, y, w-2*cm, y)
    y -= 0.7*cm

    # Beban Umum
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "Beban Umum Administrasi")
    y -= 0.5*cm
    c.setFont("Helvetica", 10)
    c.drawRightString(w-2*cm, y, f"Rp {total_beban_umum:,.0f}")
    y -= 0.5*cm
    c.line(2*cm, y, w-2*cm, y)
    y -= 0.7*cm

    # Pendapatan Luar Usaha
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "Pendapatan Luar Usaha")
    y -= 0.5*cm
    c.setFont("Helvetica", 10)
    c.drawRightString(w-2*cm, y, f"Rp {total_pendapatan_luar:,.0f}")
    y -= 0.5*cm
    c.line(2*cm, y, w-2*cm, y)
    y -= 0.7*cm

    # Beban Luar Usaha
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "Beban Luar Usaha")
    y -= 0.5*cm
    c.setFont("Helvetica", 10)
    c.drawRightString(w-2*cm, y, f"Rp {total_beban_luar:,.0f}")
    y -= 0.5*cm
    c.line(2*cm, y, w-2*cm, y)
    y -= 1*cm

    # Laba Bersih
    c.setFont("Helvetica-Bold", 11)
    c.drawString(2*cm, y, "LABA (RUGI) BERSIH")
    c.drawRightString(w-2*cm, y, f"Rp {laba_bersih:,.0f}")
    y -= 0.5*cm
    c.setLineWidth(1.2)
    c.line(w-5*cm, y, w-2*cm, y)  # garis penutup
    y -= 2*cm

    c.save()
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

st.download_button(
    "📄 Export Laba Rugi ke PDF",
    data=export_pdf_laba_rugi(),
    file_name="Laporan_Laba_Rugi.pdf",
    mime="application/pdf",
)
