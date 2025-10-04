import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

# ===== Fungsi Hitung Saldo =====
def hitung_saldo(saldo_awal, debit, kredit, posisi):
    if posisi.lower() == "debit":
        return saldo_awal + debit - kredit
    else:  # kredit
        return saldo_awal - debit + kredit

# ===== Export PDF Laba Rugi =====
def export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 3*cm

    # Judul
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 0.7*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, y, "LAPORAN LABA RUGI")
    y -= 0.6*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, f"Untuk Periode yang Berakhir Pada {periode_text}")
    y -= 1*cm

    c.setFont("Helvetica", 9)

    # Fungsi tulis baris
    def tulis_baris(text, amount=None, bold=False, total=False):
        nonlocal y
        if bold: 
            c.setFont("Helvetica-Bold", 9)
        else:
            c.setFont("Helvetica", 9)
        c.drawString(2*cm, y, text)
        if amount is not None:
            if total:
                c.line(13*cm, y-2, 19*cm, y-2)
            c.drawRightString(w-2*cm, y, f"Rp {amount:,.0f}")
        y -= 0.5*cm

    # Tampilkan detail akun
    for _, r in df_laba.iterrows():
        if r["tipe_akun"].lower() == "header":
            tulis_baris(r["nama_akun"], None, bold=True)
        else:
            tulis_baris("   " + r["nama_akun"], r["saldo_akhir_adj"])

    # Total laba bersih
    y -= 0.3*cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "LABA (RUGI) BERSIH")
    c.drawRightString(w-2*cm, y, f"Rp {laba_bersih:,.0f}")
    c.line(13*cm, y-2, 19*cm, y-2)
    c.line(13*cm, y-6, 19*cm, y-6)

    c.save()
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# ===== Export PDF Neraca =====
def export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 3*cm

    # Judul
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 0.7*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, y, "LAPORAN POSISI KEUANGAN")
    y -= 0.6*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, f"Per {periode_text}")
    y -= 1*cm

    c.setFont("Helvetica", 9)

    def tulis_baris(text, amount=None, bold=False, total=False):
        nonlocal y
        if bold: 
            c.setFont("Helvetica-Bold", 9)
        else:
            c.setFont("Helvetica", 9)
        c.drawString(2*cm, y, text)
        if amount is not None:
            if total:
                c.line(13*cm, y-2, 19*cm, y-2)
            c.drawRightString(w-2*cm, y, f"Rp {amount:,.0f}")
        y -= 0.5*cm

    # Aset
    tulis_baris("ASET", bold=True)
    for _, r in df_aset.iterrows():
        tulis_baris("   " + r["nama_akun"], r["saldo_akhir_adj"])
    tulis_baris("TOTAL ASET", total_aset, bold=True, total=True)
    y -= 0.5*cm

    # Kewajiban
    tulis_baris("KEWAJIBAN", bold=True)
    for _, r in df_kewajiban.iterrows():
        tulis_baris("   " + r["nama_akun"], r["saldo_akhir_adj"])
    tulis_baris("TOTAL KEWAJIBAN", total_kewajiban, bold=True, total=True)
    y -= 0.5*cm

    # Ekuitas
    tulis_baris("EKUITAS", bold=True)
    for _, r in df_ekuitas.iterrows():
        tulis_baris("   " + r["nama_akun"], r["saldo_akhir_adj"])
    tulis_baris("TOTAL EKUITAS", total_ekuitas, bold=True, total=True)
    y -= 0.5*cm

    # Total
    tulis_baris("TOTAL KEWAJIBAN + EKUITAS", total_kewajiban+total_ekuitas, bold=True, total=True)

    c.save()
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# ===== Streamlit App =====
st.title("📊 Laporan Keuangan")

coa_file = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])

nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
tanggal_awal = st.date_input("Tanggal Awal")
tanggal_akhir = st.date_input("Tanggal Akhir")

if coa_file and saldo_file and jurnal_file:
    coa = pd.read_excel(coa_file)
    saldo_awal = pd.read_excel(saldo_file)
    jurnal = pd.read_excel(jurnal_file)

    # Normalisasi kolom
    coa.columns = coa.columns.str.lower()
    saldo_awal.columns = saldo_awal.columns.str.lower()
    jurnal.columns = jurnal.columns.str.lower()

    if "saldo_awal" not in saldo_awal.columns:
        saldo_awal["saldo_awal"] = 0

    # Gabung data
    df = coa.merge(saldo_awal, on="kode_akun", how="left").fillna(0)
    debit = jurnal.groupby("kode_akun")["debit"].sum().reset_index()
    kredit = jurnal.groupby("kode_akun")["kredit"].sum().reset_index()
    df = df.merge(debit, on="kode_akun", how="left").merge(kredit, on="kode_akun", how="left").fillna(0)

    # Hitung saldo akhir
    df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(r["saldo_awal"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1)

    # Sesuaikan saldo adj (aturan normal)
    def adjust_saldo(r):
        if r["laporan"] == "Laporan Posisi Keuangan":
            if "aset" in r["sub_tipe_laporan"].lower() and r["posisi_normal_akun"].lower() == "debit":
                return r["saldo_akhir"]
            elif "kewajiban" in r["sub_tipe_laporan"].lower() and r["posisi_normal_akun"].lower() == "kredit":
                return r["saldo_akhir"]
            elif "ekuitas" in r["sub_tipe_laporan"].lower() and r["posisi_normal_akun"].lower() == "kredit":
                return r["saldo_akhir"]
            else:
                return -r["saldo_akhir"]
        else:
            return r["saldo_akhir"]

    df["saldo_akhir_adj"] = df.apply(adjust_saldo, axis=1)

    # Filter laporan
    df_laba = df[df["laporan"]=="Laporan Laba Rugi"]
    df_aset = df[(df["laporan"]=="Laporan Posisi Keuangan") & (df["sub_tipe_laporan"].str.contains("Aset", case=False))]
    df_kewajiban = df[(df["laporan"]=="Laporan Posisi Keuangan") & (df["sub_tipe_laporan"].str.contains("Kewajiban", case=False))]
    df_ekuitas = df[(df["laporan"]=="Laporan Posisi Keuangan") & (df["sub_tipe_laporan"].str.contains("Ekuitas", case=False))]

    total_pendapatan = df_laba[df_laba["sub_tipe_laporan"].str.contains("Pendapatan", case=False)]["saldo_akhir_adj"].sum()
    total_beban = df_laba[df_laba["sub_tipe_laporan"].str.contains("Beban", case=False)]["saldo_akhir_adj"].sum()
    laba_bersih = total_pendapatan - total_beban

    # Tambahkan laba bersih ke ekuitas
    df_ekuitas = pd.concat([df_ekuitas, pd.DataFrame([{"kode_akun":"3004","nama_akun":"Laba (Rugi) Berjalan","saldo_akhir_adj":laba_bersih}])])

    total_aset = df_aset["saldo_akhir_adj"].sum()
    total_kewajiban = df_kewajiban["saldo_akhir_adj"].sum()
    total_ekuitas = df_ekuitas["saldo_akhir_adj"].sum()

    periode_text = tanggal_akhir.strftime("%d %B %Y")

    # PDF Laba Rugi
    pdf_laba = export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text)
    st.download_button("⬇️ Download Laba Rugi (PDF)", pdf_laba, file_name="Laporan_Laba_Rugi.pdf")

    # PDF Neraca
    pdf_nr = export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text)
    st.download_button("⬇️ Download Neraca (PDF)", pdf_nr, file_name="Laporan_Posisi_Keuangan.pdf")

    # Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_laba.to_excel(writer, sheet_name="Laba Rugi", index=False)
        df_aset.to_excel(writer, sheet_name="Aset", index=False)
        df_kewajiban.to_excel(writer, sheet_name="Kewajiban", index=False)
        df_ekuitas.to_excel(writer, sheet_name="Ekuitas", index=False)
        coa.to_excel(writer, sheet_name="COA", index=False)
        saldo_awal.to_excel(writer, sheet_name="Saldo Awal", index=False)
        jurnal.to_excel(writer, sheet_name="Jurnal", index=False)
    st.download_button("⬇️ Download Excel", output.getvalue(), file_name="Laporan_Keuangan.xlsx")
