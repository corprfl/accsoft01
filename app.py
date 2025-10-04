import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import datetime

# =============================
# Fungsi Hitung Saldo Akhir
# =============================
def hitung_saldo(saldo_awal, debit, kredit, posisi_normal):
    if posisi_normal.lower() == "debit":
        return saldo_awal + debit - kredit
    elif posisi_normal.lower() == "kredit":
        return saldo_awal - debit + kredit
    else:
        return saldo_awal + debit - kredit

# =============================
# Fungsi Adjust Saldo Neraca
# =============================
def adjust_neraca_value(row):
    if "aset" in str(row["sub_tipe_laporan"]).lower():
        return row["saldo_akhir"] if row["posisi_normal_akun"].lower()=="debit" else -row["saldo_akhir"]
    elif "kewajiban" in str(row["sub_tipe_laporan"]).lower():
        return row["saldo_akhir"] if row["posisi_normal_akun"].lower()=="kredit" else -row["saldo_akhir"]
    elif "ekuitas" in str(row["sub_tipe_laporan"]).lower():
        return row["saldo_akhir"] if row["posisi_normal_akun"].lower()=="kredit" else -row["saldo_akhir"]
    else:
        return row["saldo_akhir"]

# =============================
# Export PDF Laba Rugi
# =============================
def export_pdf_laba_rugi(df, laba_bersih, nama_pt, periode_text):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 2*cm

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 20
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, "LAPORAN LABA RUGI")
    y -= 15
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, f"Untuk Periode yang Berakhir Pada {periode_text}")
    y -= 30

    def tulis_baris(label, value=None, bold=False, garis=False, double=False):
        nonlocal y
        if bold: c.setFont("Helvetica-Bold", 10)
        else: c.setFont("Helvetica", 10)
        c.drawString(2*cm, y, str(label))
        if value is not None:
            c.drawRightString(w-2*cm, y, f"Rp {value:,.0f}")
        y -= 15
        if garis:
            c.line(w-6*cm, y+5, w-2*cm, y+5)
        if double:
            c.line(w-6*cm, y+5, w-2*cm, y+5)
            c.line(w-6*cm, y+2, w-2*cm, y+2)

    # Struktur laporan
    for section in ["Pendapatan", "Beban Umum Administrasi", "Pendapatan Luar Usaha", "Beban Luar Usaha"]:
        sub = df[df["sub_tipe_laporan"]==section]
        if not sub.empty:
            tulis_baris(section, None, bold=True)
            for _, r in sub.iterrows():
                if r["tipe_akun"].lower() == "header":  # header kosong
                    tulis_baris(r["nama_akun"], None, bold=True)
                else:
                    tulis_baris("   "+r["nama_akun"], r["saldo_akhir_adj"])
            total = sub["saldo_akhir_adj"].sum()
            tulis_baris(f"TOTAL {section.upper()}", total, bold=True, garis=True)

    # Laba bersih
    tulis_baris(f"LABA (RUGI) BERSIH", laba_bersih, bold=True, double=True)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# =============================
# Export PDF Neraca
# =============================
def export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 2*cm

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 20
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, "LAPORAN POSISI KEUANGAN")
    y -= 15
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, f"Per {periode_text}")
    y -= 30

    def draw_sec(title, data, total, double=False):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, title)
        y -= 15
        c.setFont("Helvetica", 10)
        for _, r in data.iterrows():
            if r["tipe_akun"].lower()=="header":
                c.drawString(2.3*cm, y, r["nama_akun"])
            else:
                c.drawString(2.6*cm, y, r["nama_akun"])
                c.drawRightString(w-2*cm, y, f"Rp {r['saldo_akhir_adj']:,.0f}")
            y -= 15
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, f"TOTAL {title.upper()}")
        c.drawRightString(w-2*cm, y, f"Rp {total:,.0f}")
        c.line(w-6*cm, y+5, w-2*cm, y+5)
        if double:
            c.line(w-6*cm, y+2, w-2*cm, y+2)
        y -= 25

    draw_sec("ASET", df_aset, total_aset)
    draw_sec("KEWAJIBAN", df_kewajiban, total_kewajiban)
    draw_sec("EKUITAS", df_ekuitas, total_ekuitas)
    draw_sec("KEWAJIBAN + EKUITAS", pd.DataFrame(), total_kewajiban+total_ekuitas, double=True)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# =============================
# Streamlit App
# =============================
st.title("📊 Laporan Keuangan Otomatis")

uploaded_coa = st.file_uploader("Upload COA.xlsx", type="xlsx")
uploaded_saldo = st.file_uploader("Upload Saldo Awal.xlsx", type="xlsx")
uploaded_jurnal = st.file_uploader("Upload Jurnal.xlsx", type="xlsx")

nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
pejabat = st.text_input("Pejabat Penandatangan", "Direktur")

tanggal_awal = st.date_input("Tanggal Awal", datetime.date(2025,1,1))
tanggal_akhir = st.date_input("Tanggal Akhir", datetime.date(2025,12,31))

if uploaded_coa and uploaded_saldo and uploaded_jurnal:
    coa = pd.read_excel(uploaded_coa)
    saldo_awal = pd.read_excel(uploaded_saldo)
    jurnal = pd.read_excel(uploaded_jurnal)

    # pastikan kolom numeric
    for col in ["saldo","saldo_awal","debit","kredit"]:
        if col in saldo_awal.columns:
            saldo_awal[col] = pd.to_numeric(saldo_awal[col], errors="coerce").fillna(0)
        if col in jurnal.columns:
            jurnal[col] = pd.to_numeric(jurnal[col], errors="coerce").fillna(0)

    # gabung COA + Saldo
    df = coa.merge(saldo_awal, on="kode_akun", how="left").fillna(0)

    # agregat jurnal
    agg_jurnal = jurnal.groupby("kode_akun")[["debit","kredit"]].sum().reset_index()
    df = df.merge(agg_jurnal, on="kode_akun", how="left").fillna(0)

    # hitung saldo akhir
    df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(r.get("saldo",0), r.get("debit",0), r.get("kredit",0), r["posisi_normal_akun"]), axis=1)
    df["saldo_akhir_adj"] = df.apply(adjust_neraca_value, axis=1)

    # Laba Rugi
    df_laba = df[df["laporan"].str.contains("Laba Rugi", case=False, na=False)]
    laba_bersih = df_laba[df_laba["sub_tipe_laporan"].isin(["Pendapatan","Pendapatan Luar Usaha"])]["saldo_akhir_adj"].sum() \
                 - df_laba[df_laba["sub_tipe_laporan"].isin(["Beban Umum Administrasi","Beban Luar Usaha"])]["saldo_akhir_adj"].sum()

    # Neraca
    df_aset = df[df["sub_tipe_laporan"].str.contains("Aset",case=False,na=False)]
    df_kewajiban = df[df["sub_tipe_laporan"].str.contains("Kewajiban",case=False,na=False)]
    df_ekuitas = df[df["sub_tipe_laporan"].str.contains("Ekuitas",case=False,na=False)].copy()

    # Update laba berjalan
    if "3004" in df_ekuitas["kode_akun"].values:
        df_ekuitas.loc[df_ekuitas["kode_akun"]=="3004","saldo_akhir_adj"] = laba_bersih

    total_aset = df_aset["saldo_akhir_adj"].sum()
    total_kewajiban = df_kewajiban["saldo_akhir_adj"].sum()
    total_ekuitas = df_ekuitas["saldo_akhir_adj"].sum()

    periode_text = tanggal_akhir.strftime("%d %B %Y")

    # Export PDF
    pdf_laba = export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text)
    pdf_neraca = export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text)

    st.download_button("📥 Download Laporan Laba Rugi (PDF)", pdf_laba, file_name="Laba_Rugi.pdf")
    st.download_button("📥 Download Laporan Posisi Keuangan (PDF)", pdf_neraca, file_name="Neraca.pdf")

    st.success("✅ Laporan berhasil dibuat.")
