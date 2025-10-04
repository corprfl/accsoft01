import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

st.set_page_config(page_title="Laporan Keuangan", layout="wide")

st.title("📊 Aplikasi Laporan Keuangan")

# ==============================
# UPLOAD FILES
# ==============================
coa_file = st.file_uploader("Upload Chart of Account (COA)", type=["xlsx"])
saldo_file = st.file_uploader("Upload Saldo Awal", type=["xlsx"])
jurnal_file = st.file_uploader("Upload Jurnal Umum", type=["xlsx"])

tanggal_awal = st.date_input("Tanggal Awal Periode")
tanggal_akhir = st.date_input("Tanggal Akhir Periode")

nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
pejabat = st.text_input("Nama Pejabat", "Reza Fahlevi Lubis")

if coa_file and saldo_file and jurnal_file:

    # ==============================
    # BACA DATA + NORMALISASI KOLOM
    # ==============================
    coa = pd.read_excel(coa_file)
    saldo_awal = pd.read_excel(saldo_file)
    jurnal = pd.read_excel(jurnal_file)

    # Normalisasi nama kolom
    coa.columns = coa.columns.str.strip().str.lower()
    saldo_awal.columns = saldo_awal.columns.str.strip().str.lower()
    jurnal.columns = jurnal.columns.str.strip().str.lower()

    # Samakan nama kolom
    rename_map = {
        "kode akun": "kode_akun",
        "nama akun": "nama_akun",
        "posisi normal akun": "posisi_normal_akun",
        "sub tipe laporan": "sub_tipe_laporan",
        "saldo": "saldo",
        "saldo awal": "saldo",
        "debit": "debit",
        "kredit": "kredit"
    }
    coa.rename(columns=rename_map, inplace=True)
    saldo_awal.rename(columns=rename_map, inplace=True)
    jurnal.rename(columns=rename_map, inplace=True)

    # Pastikan kolom saldo ada
    if "saldo" not in saldo_awal.columns:
        saldo_awal["saldo"] = 0

    # ==============================
    # GABUNG DATA
    # ==============================
    df = coa.merge(saldo_awal[["kode_akun","saldo"]], on="kode_akun", how="left").fillna(0)
    jurnal_sum = jurnal.groupby("kode_akun").agg({"debit":"sum","kredit":"sum"}).reset_index()
    df = df.merge(jurnal_sum, on="kode_akun", how="left").fillna(0)

    # ==============================
    # HITUNG SALDO AKHIR
    # ==============================
    def hitung_saldo(saldo_awal, debit, kredit, normal):
        if str(normal).lower() == "debit":
            return saldo_awal + debit - kredit
        else:
            return saldo_awal - debit + kredit

    df["saldo_akhir"] = df.apply(
        lambda r: hitung_saldo(r["saldo"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1
    )

    # ==============================
    # RULE PENYESUAIAN LABA RUGI
    # ==============================
    def adjust_laba(row):
        if "pendapatan" in str(row["sub_tipe_laporan"]).lower():
            return abs(row["saldo_akhir"])
        elif "beban" in str(row["sub_tipe_laporan"]).lower():
            return abs(row["saldo_akhir"])
        else:
            return row["saldo_akhir"]

    df["saldo_lr"] = df.apply(adjust_laba, axis=1)

    # ==============================
    # RULE PENYESUAIAN NERACA
    # ==============================
    def adjust_neraca(row):
        if "aset" in str(row["sub_tipe_laporan"]).lower():
            return row["saldo_akhir"] if str(row["posisi_normal_akun"]).lower()=="debit" else -row["saldo_akhir"]
        elif "kewajiban" in str(row["sub_tipe_laporan"]).lower():
            return row["saldo_akhir"] if str(row["posisi_normal_akun"]).lower()=="kredit" else -row["saldo_akhir"]
        elif "ekuitas" in str(row["sub_tipe_laporan"]).lower():
            return row["saldo_akhir"] if str(row["posisi_normal_akun"]).lower()=="kredit" else -row["saldo_akhir"]
        else:
            return row["saldo_akhir"]

    df["saldo_nr"] = df.apply(adjust_neraca, axis=1)

    # ==============================
    # LAPORAN LABA RUGI
    # ==============================
    df_laba = df[df["laporan"].str.contains("Laba Rugi", case=False, na=False)]

    total_pendapatan = df_laba[df_laba["sub_tipe_laporan"].str.contains("pendapatan", case=False, na=False)]["saldo_lr"].sum()
    total_beban = df_laba[df_laba["sub_tipe_laporan"].str.contains("beban", case=False, na=False)]["saldo_lr"].sum()
    laba_bersih = total_pendapatan - total_beban

    # ==============================
    # LAPORAN NERACA
    # ==============================
    df_aset = df[df["sub_tipe_laporan"].str.contains("aset", case=False, na=False)]
    df_kewajiban = df[df["sub_tipe_laporan"].str.contains("kewajiban", case=False, na=False)]
    df_ekuitas = df[df["sub_tipe_laporan"].str.contains("ekuitas", case=False, na=False)].copy()

    # isi laba berjalan dari laba rugi
    df_ekuitas.loc[df_ekuitas["kode_akun"]=="3004", "saldo_nr"] = laba_bersih

    total_aset = df_aset["saldo_nr"].sum()
    total_kewajiban = df_kewajiban["saldo_nr"].sum()
    total_ekuitas = df_ekuitas["saldo_nr"].sum()

    # ==============================
    # EXPORT PDF FUNKSI
    # ==============================
    def export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode):
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        w, h = A4
        y = h - 2*cm
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(w/2, y, nama_pt)
        y -= 20
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(w/2, y, "LAPORAN LABA RUGI")
        y -= 15
        c.setFont("Helvetica", 10)
        c.drawCentredString(w/2, y, f"Untuk Periode yang Berakhir Pada {periode}")

        y -= 40
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, "TOTAL PENDAPATAN")
        c.drawRightString(w-2*cm, y, f"Rp {total_pendapatan:,.0f}")

        y -= 20
        c.drawString(2*cm, y, "TOTAL BEBAN")
        c.drawRightString(w-2*cm, y, f"Rp {total_beban:,.0f}")

        y -= 30
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2*cm, y, "LABA (RUGI) BERSIH")
        c.drawRightString(w-2*cm, y, f"Rp {laba_bersih:,.0f}")
        c.line(w-6*cm, y-2, w-2*cm, y-2)
        c.line(w-6*cm, y-6, w-2*cm, y-6)

        c.showPage()
        c.save()
        buffer.seek(0)
        return buffer

    def export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode):
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        w, h = A4
        y = h - 2*cm
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(w/2, y, nama_pt)
        y -= 20
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(w/2, y, "LAPORAN POSISI KEUANGAN")
        y -= 15
        c.setFont("Helvetica", 10)
        c.drawCentredString(w/2, y, f"Per {periode}")

        y -= 40
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, "TOTAL ASET")
        c.drawRightString(w-2*cm, y, f"Rp {total_aset:,.0f}")

        y -= 20
        c.drawString(2*cm, y, "TOTAL KEWAJIBAN")
        c.drawRightString(w-2*cm, y, f"Rp {total_kewajiban:,.0f}")

        y -= 20
        c.drawString(2*cm, y, "TOTAL EKUITAS")
        c.drawRightString(w-2*cm, y, f"Rp {total_ekuitas:,.0f}")

        y -= 30
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2*cm, y, "TOTAL KEWAJIBAN + EKUITAS")
        c.drawRightString(w-2*cm, y, f"Rp {total_kewajiban+total_ekuitas:,.0f}")
        c.line(w-8*cm, y-2, w-2*cm, y-2)
        c.line(w-8*cm, y-6, w-2*cm, y-6)

        c.showPage()
        c.save()
        buffer.seek(0)
        return buffer

    # ==============================
    # TAMPILKAN & DOWNLOAD
    # ==============================
    st.subheader("📑 Laporan Laba Rugi")
    st.metric("LABA (RUGI) BERSIH", f"Rp {laba_bersih:,.0f}")

    pdf_lr = export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, tanggal_akhir.strftime("%d %B %Y"))
    st.download_button("⬇️ Download Laba Rugi (PDF)", data=pdf_lr, file_name="Laporan_Laba_Rugi.pdf")

    st.subheader("📑 Laporan Posisi Keuangan (Neraca)")
    st.metric("TOTAL ASET", f"Rp {total_aset:,.0f}")
    st.metric("TOTAL KEWAJIBAN + EKUITAS", f"Rp {total_kewajiban+total_ekuitas:,.0f}")

    pdf_nr = export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, tanggal_akhir.strftime("%d %B %Y"))
    st.download_button("⬇️ Download Neraca (PDF)", data=pdf_nr, file_name="Laporan_Neraca.pdf")
