import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io

# =========================
# Fungsi Export Laba Rugi
# =========================
def export_pdf_labarugi(df_laba, total_pendapatan, total_beban, total_luar, laba_rugi, nama_pt, periode_text):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    # Judul
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, h-2*cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, h-2.7*cm, "LAPORAN LABA RUGI")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, h-3.3*cm, f"Untuk Periode yang Berakhir Pada {periode_text}")

    # Border box
    margin_x = 2*cm
    top_y = h-3.6*cm
    bottom_y = 3*cm
    c.rect(margin_x-0.5*cm, bottom_y-0.5*cm, w-2*(margin_x-0.5*cm), top_y-bottom_y+0.5*cm)

    y = top_y - 1*cm
    line_height = 14

    # Fungsi nulis baris akun
    def tulis_baris(label, amount=None, bold=False):
        nonlocal y
        if bold:
            c.setFont("Helvetica-Bold", 10)
        else:
            c.setFont("Helvetica", 10)

        c.drawString(margin_x, y, str(label))
        if amount is not None:
            c.drawRightString(w-margin_x, y, f"Rp {amount:,.0f}")

        y -= line_height

    # Fungsi total dengan underline
    def tulis_total(label, amount):
        nonlocal y
        c.line(margin_x, y+3, w-margin_x, y+3)   # garis atas
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margin_x, y, str(label))
        c.drawRightString(w-margin_x, y, f"Rp {amount:,.0f}")
        y -= line_height
        c.line(margin_x, y+line_height-3, w-margin_x, y+line_height-3)  # garis bawah

    # =====================
    # Isi Laporan
    # =====================
    tulis_baris("Pendapatan", bold=True)
    for _, r in df_laba[df_laba['sub_tipe_laporan']=="Pendapatan"].iterrows():
        tulis_baris("   " + r['nama_akun'], r['saldo_akhir_adj'])
    tulis_total("TOTAL PENDAPATAN", total_pendapatan)
    y -= line_height

    tulis_baris("Beban Umum Administrasi", bold=True)
    for _, r in df_laba[df_laba['sub_tipe_laporan']=="Beban Umum Administrasi"].iterrows():
        tulis_baris("   " + r['nama_akun'], r['saldo_akhir_adj'])
    tulis_total("TOTAL BEBAN UMUM ADMINISTRASI", total_beban)
    y -= line_height

    tulis_baris("Pendapatan Luar Usaha", bold=True)
    for _, r in df_laba[df_laba['sub_tipe_laporan']=="Pendapatan Luar Usaha"].iterrows():
        tulis_baris("   " + r['nama_akun'], r['saldo_akhir_adj'])
    tulis_total("TOTAL PENDAPATAN LUAR USAHA", total_luar)
    y -= line_height

    tulis_baris("Beban Luar Usaha", bold=True)
    for _, r in df_laba[df_laba['sub_tipe_laporan']=="Beban Luar Usaha"].iterrows():
        tulis_baris("   " + r['nama_akun'], r['saldo_akhir_adj'])
    tulis_total("TOTAL BEBAN LUAR USAHA", df_laba[df_laba['sub_tipe_laporan']=="Beban Luar Usaha"]['saldo_akhir_adj'].sum())
    y -= line_height*2

    # Laba (Rugi) Bersih dengan double underline
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin_x, y, "LABA (RUGI) BERSIH")
    c.drawRightString(w-margin_x, y, f"Rp {laba_rugi:,.0f}")
    c.line(w-margin_x-200, y-2, w-margin_x, y-2)
    c.line(w-margin_x-200, y-5, w-margin_x, y-5)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Laporan Laba Rugi", layout="wide")

st.title("📊 Laporan Laba Rugi")

# Upload contoh data
coa = pd.read_excel("COA.xlsx")
saldo_awal = pd.read_excel("Saldo Awal.xlsx")
jurnal = pd.read_excel("Jurnal.xlsx")

# Dummy hasil hitungan (contoh)
df_laba = pd.DataFrame({
    "nama_akun": ["Pendapatan", "Biaya Gaji", "Biaya ATK", "Pendapatan Bunga", "Beban Lain-lain"],
    "sub_tipe_laporan": ["Pendapatan", "Beban Umum Administrasi", "Beban Umum Administrasi", "Pendapatan Luar Usaha", "Beban Luar Usaha"],
    "saldo_akhir_adj": [8731100054, 4000000000, 11851000, 83750197, 4552826]
})

total_pendapatan = 8731100054
total_beban = 6102136416
total_luar = 83750197
total_beban_luar = 4552826
laba_rugi = total_pendapatan + total_luar - total_beban - total_beban_luar

nama_pt = "PT Contoh Sejahtera"
periode_text = "31 Desember 2025"

if st.button("📄 Export PDF Laba Rugi"):
    pdf_buf = export_pdf_labarugi(df_laba, total_pendapatan, total_beban, total_luar, laba_rugi, nama_pt, periode_text)
    st.download_button("Download Laba Rugi PDF", data=pdf_buf, file_name="Laporan_Laba_Rugi.pdf", mime="application/pdf")
