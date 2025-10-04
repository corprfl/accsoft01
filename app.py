import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from datetime import datetime

# ==============================
# FUNGSI PERHITUNGAN SALDO
# ==============================
def hitung_saldo(saldo_awal, debit, kredit, posisi_normal):
    if posisi_normal.lower() == "debit":
        return saldo_awal + debit - kredit
    else:
        return saldo_awal - debit + kredit

# ==============================
# FUNGSI EXPORT PDF LABA RUGI
# ==============================
def export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 2 * cm

    # Judul
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 0.8*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, y, "LAPORAN LABA RUGI")
    y -= 0.6*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, f"Untuk Periode yang Berakhir Pada {periode_text}")
    y -= 1*cm

    # Kotak Border
    c.rect(2*cm, 2*cm, w-4*cm, y-2*cm)

    def tulis_header(teks):
        nonlocal y
        if y < 3*cm: 
            c.showPage(); y = h-2*cm
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2.2*cm, y, teks)
        y -= 0.5*cm

    def tulis_baris(teks, nilai):
        nonlocal y
        if y < 3*cm:
            c.showPage(); y = h-2*cm
        c.setFont("Helvetica", 9)
        c.drawString(3*cm, y, teks)
        if nilai != 0:
            c.drawRightString(w-2.2*cm, y, f"Rp {nilai:,.0f}")
        y -= 0.4*cm

    def tulis_total(teks, nilai):
        nonlocal y
        if y < 3*cm:
            c.showPage(); y = h-2*cm
        c.setFont("Helvetica-Bold", 9)
        # garis atas penjumlahan
        c.line(w-5*cm, y+0.2*cm, w-2*cm, y+0.2*cm)
        c.drawString(2.5*cm, y, teks.upper())
        c.drawRightString(w-2.2*cm, y, f"Rp {nilai:,.0f}")
        y -= 0.5*cm

    # ========== CETAK LAPORAN ==========
    sub_tipe = None
    subtotal = 0
    for _, r in df_laba.iterrows():
        if str(r["tipe_akun"]).lower() == "header":
            # cetak total sub sebelumnya
            if sub_tipe and subtotal != 0:
                tulis_total(f"TOTAL {sub_tipe}", subtotal)
                subtotal = 0
            sub_tipe = r["nama_akun"]
            tulis_header(sub_tipe)
        else:
            if r["saldo_akhir"] != 0:
                tulis_baris(r["nama_akun"], r["saldo_akhir"])
                subtotal += r["saldo_akhir"]

    # total terakhir
    if sub_tipe and subtotal != 0:
        tulis_total(f"TOTAL {sub_tipe}", subtotal)

    # Laba Rugi Bersih
    y -= 0.5*cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "LABA (RUGI) BERSIH")
    # garis double
    c.line(w-5*cm, y+0.2*cm, w-2*cm, y+0.2*cm)
    c.line(w-5*cm, y+0.35*cm, w-2*cm, y+0.35*cm)
    c.drawRightString(w-2.2*cm, y, f"Rp {laba_bersih:,.0f}")
    y -= 1*cm

    c.save()
    buffer.seek(0)
    return buffer

# ==============================
# STREAMLIT APP
# ==============================
st.title("📊 Laporan Keuangan")

uploaded_coa = st.file_uploader("Upload COA.xlsx", type="xlsx")
uploaded_saldo = st.file_uploader("Upload Saldo Awal.xlsx", type="xlsx")
uploaded_jurnal = st.file_uploader("Upload Jurnal.xlsx", type="xlsx")

nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
periode_akhir = st.date_input("Tanggal Akhir Periode", datetime.today())
periode_text = periode_akhir.strftime("%d %B %Y")

if uploaded_coa and uploaded_saldo and uploaded_jurnal:
    coa = pd.read_excel(uploaded_coa)
    saldo_awal = pd.read_excel(uploaded_saldo)
    jurnal = pd.read_excel(uploaded_jurnal)

    # pastikan nama kolom lower case
    coa.columns = coa.columns.str.lower()
    saldo_awal.columns = saldo_awal.columns.str.lower()
    jurnal.columns = jurnal.columns.str.lower()

    # merge saldo awal + jurnal
    df = coa.merge(saldo_awal, on="kode_akun", how="left").fillna(0)
    jurnal_group = jurnal.groupby("kode_akun").agg({"debit":"sum","kredit":"sum"}).reset_index()
    df = df.merge(jurnal_group, on="kode_akun", how="left").fillna(0)

    df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(
        r["saldo"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1)

    # Filter Laba Rugi
    df_laba = df[df["laporan"].str.contains("laba", case=False)]
    laba_bersih = df_laba[df_laba["tipe_akun"]!="Header"]["saldo_akhir"].sum()

    st.subheader("Laporan Laba Rugi")
    st.metric("LABA (RUGI) BERSIH", f"Rp {laba_bersih:,.0f}")

    if st.button("📄 Export Laba Rugi ke PDF"):
        pdf_lr = export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text)
        st.download_button("Download PDF", data=pdf_lr, file_name="Laba_Rugi.pdf", mime="application/pdf")
