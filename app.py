import streamlit as st
import pandas as pd
import io
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

st.set_page_config(page_title="Laporan Keuangan", layout="wide")

st.title("📊 Generator Laporan Keuangan")

uploaded_coa = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
uploaded_saldo = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
uploaded_jurnal = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])

tanggal_awal = st.date_input("Tanggal Awal Periode")
tanggal_akhir = st.date_input("Tanggal Akhir Periode")

def hitung_saldo(saldo, debit, kredit, normal):
    if normal.lower() == "debit":
        return saldo + debit - kredit
    else:
        return saldo - debit + kredit

def buat_pdf_laba_rugi(df, laba_bersih, nama_pt, periode_text):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, h-2*cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, h-3*cm, "LAPORAN LABA RUGI")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, h-3.7*cm, f"Untuk Periode yang Berakhir Pada {periode_text}")

    y = h-5*cm
    c.setFont("Helvetica", 10)

    def tulis(header, amount=None, total=False):
        nonlocal y
        if y < 3*cm:
            c.showPage()
            y = h-2*cm
        if amount is None:
            c.setFont("Helvetica-Bold", 10)
            c.drawString(3*cm, y, header)
        else:
            if total:
                c.line(14*cm, y+3, w-2*cm, y+3)
                c.setFont("Helvetica-Bold", 10)
            else:
                c.setFont("Helvetica", 10)
            c.drawString(3*cm, y, header)
            c.drawRightString(w-2*cm, y, f"Rp {amount:,.0f}")
        y -= 0.5*cm

    for _, r in df.iterrows():
        if "header" in str(r["tipe_akun"]).lower():
            tulis(r["nama_akun"])
        else:
            if r["saldo_akhir_adj"] != 0:
                tulis("   "+r["nama_akun"], r["saldo_akhir_adj"])

    c.line(14*cm, y, w-2*cm, y)
    c.line(14*cm, y-3, w-2*cm, y-3)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(3*cm, y-1*cm, "LABA (RUGI) BERSIH")
    c.drawRightString(w-2*cm, y-1*cm, f"Rp {laba_bersih:,.0f}")

    c.rect(2*cm, 2*cm, w-4*cm, h-6*cm)

    c.save()
    buffer.seek(0)
    return buffer

def buat_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, h-2*cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, h-3*cm, "LAPORAN POSISI KEUANGAN")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, h-3.7*cm, f"Per {periode_text}")

    y = h-5*cm

    def draw_sec(judul, df, total):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2.5*cm, y, judul)
        y -= 0.5*cm
        for _, r in df.iterrows():
            if r["saldo_akhir_adj"] != 0:
                c.setFont("Helvetica", 10)
                c.drawString(3*cm, y, r["nama_akun"])
                c.drawRightString(w-2*cm, y, f"Rp {r['saldo_akhir_adj']:,.0f}")
                y -= 0.4*cm
        c.line(14*cm, y, w-2*cm, y)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(3*cm, y-0.5*cm, f"TOTAL {judul}")
        c.drawRightString(w-2*cm, y-0.5*cm, f"Rp {total:,.0f}")
        y -= 1*cm

    draw_sec("ASET", df_aset, total_aset)
    draw_sec("KEWAJIBAN", df_kewajiban, total_kewajiban)
    draw_sec("EKUITAS", df_ekuitas, total_ekuitas)

    c.line(14*cm, y, w-2*cm, y)
    c.line(14*cm, y-3, w-2*cm, y-3)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(3*cm, y-1*cm, "TOTAL KEWAJIBAN + EKUITAS")
    c.drawRightString(w-2*cm, y-1*cm, f"Rp {(total_kewajiban+total_ekuitas):,.0f}")

    c.rect(2*cm, 2*cm, w-4*cm, h-6*cm)

    c.save()
    buffer.seek(0)
    return buffer

if uploaded_coa and uploaded_saldo and uploaded_jurnal:
    coa = pd.read_excel(uploaded_coa)
    saldo_awal = pd.read_excel(uploaded_saldo)
    jurnal = pd.read_excel(uploaded_jurnal)

    if "saldo" in saldo_awal.columns:
        saldo_awal.rename(columns={"saldo":"saldo_awal"}, inplace=True)

    df = coa.merge(saldo_awal, on="kode_akun", how="left").fillna(0)
    jurnal["tanggal"] = pd.to_datetime(jurnal["tanggal"])

    mutasi = jurnal[(jurnal["tanggal"]>=pd.to_datetime(tanggal_awal)) & (jurnal["tanggal"]<=pd.to_datetime(tanggal_akhir))]
    mutasi_group = mutasi.groupby("kode_akun").agg({"debit":"sum","kredit":"sum"}).reset_index()
    df = df.merge(mutasi_group, on="kode_akun", how="left").fillna(0)

    df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(r["saldo_awal"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1)

    df["saldo_akhir_adj"] = df.apply(lambda r: r["saldo_akhir"] if r["posisi_normal_akun"].lower()=="debit" else -r["saldo_akhir"] if r["saldo_akhir"]<0 else r["saldo_akhir"], axis=1)

    df_laba = df[df["laporan"]=="Laporan Laba Rugi"]
    df_neraca = df[df["laporan"]=="Laporan Posisi Keuangan"]

    laba_bersih = df_laba["saldo_akhir_adj"].sum()
    df_aset = df_neraca[df_neraca["sub_tipe_laporan"]=="Aset Lancar"]
    df_kewajiban = df_neraca[df_neraca["sub_tipe_laporan"]=="Kewajiban"]
    df_ekuitas = df_neraca[df_neraca["sub_tipe_laporan"]=="Ekuitas"].copy()
    df_ekuitas.loc[df_ekuitas["nama_akun"].str.contains("Laba", case=False), "saldo_akhir_adj"] = laba_bersih

    total_aset = df_aset["saldo_akhir_adj"].sum()
    total_kewajiban = df_kewajiban["saldo_akhir_adj"].sum()
    total_ekuitas = df_ekuitas["saldo_akhir_adj"].sum()

    periode_text = tanggal_akhir.strftime("%d %B %Y")

    pdf_lr = buat_pdf_laba_rugi(df_laba, laba_bersih, "PT Contoh Sejahtera", periode_text)
    pdf_nr = buat_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, "PT Contoh Sejahtera", periode_text)

    st.download_button("⬇️ Download Laporan Laba Rugi (PDF)", data=pdf_lr, file_name="Laporan_Laba_Rugi.pdf")
    st.download_button("⬇️ Download Laporan Posisi Keuangan (PDF)", data=pdf_nr, file_name="Laporan_Posisi_Keuangan.pdf")
