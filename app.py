import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

# -----------------------------
# Fungsi hitung saldo akhir
# -----------------------------
def hitung_saldo(saldo_awal, debit, kredit, posisi_normal):
    if posisi_normal.lower() == "debit":
        return saldo_awal + debit - kredit
    else:  # kredit
        return saldo_awal - debit + kredit

# -----------------------------
# Fungsi export PDF
# -----------------------------
def export_pdf(df_laba_rugi, laba_rugi, df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, pejabat):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Judul
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, height-2*cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, height-2.7*cm, "LAPORAN LABA RUGI & NERACA")
    c.setFont("Helvetica", 10)
    c.drawCentredString(width/2, height-3.3*cm, "Periode: 31 Desember 2025")

    y = height-4*cm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(2*cm, y, "LAPORAN LABA RUGI")
    y -= 0.5*cm

    # Tampilkan per bagian Laba Rugi
    for laporan, df in df_laba_rugi.items():
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, laporan)
        y -= 0.4*cm
        c.setFont("Helvetica", 9)
        for _, row in df.iterrows():
            c.drawString(2.2*cm, y, str(row['nama_akun']))
            c.drawRightString(width-2*cm, y, f"Rp {row['saldo_akhir']:,.0f}")
            y -= 0.35*cm
        c.setFont("Helvetica-Bold", 9)
        c.drawString(2*cm, y, f"TOTAL {laporan.upper()}")
        c.drawRightString(width-2*cm, y, f"Rp {df['saldo_akhir'].sum():,.0f}")
        y -= 0.5*cm

    c.setFont("Helvetica-Bold", 11)
    c.drawString(2*cm, y, f"LABA (RUGI) BERSIH : Rp {laba_rugi:,.0f}")
    y -= 1*cm

    # Neraca
    c.setFont("Helvetica-Bold", 11)
    c.drawString(2*cm, y, "LAPORAN POSISI KEUANGAN (NERACA)")
    y -= 0.5*cm

    def draw_section(title, df, total):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, title)
        y -= 0.4*cm
        c.setFont("Helvetica", 9)
        for _, row in df.iterrows():
            c.drawString(2.2*cm, y, str(row['nama_akun']))
            c.drawRightString(width-2*cm, y, f"Rp {row['saldo_akhir']:,.0f}")
            y -= 0.35*cm
        c.setFont("Helvetica-Bold", 9)
        c.drawString(2*cm, y, f"TOTAL {title.upper()} :")
        c.drawRightString(width-2*cm, y, f"Rp {total:,.0f}")
        y -= 0.6*cm

    draw_section("ASET", df_aset, total_aset)
    draw_section("KEWAJIBAN", df_kewajiban, total_kewajiban)
    draw_section("EKUITAS", df_ekuitas, total_ekuitas)

    # tanda tangan
    c.setFont("Helvetica", 10)
    c.drawString(width-8*cm, 3*cm, "Direktur,")
    c.drawString(width-8*cm, 2*cm, pejabat)

    c.save()
    buffer.seek(0)
    return buffer

# -----------------------------
# App Streamlit
# -----------------------------
st.title("📊 Laporan Keuangan Generator")

coa_file = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])

nama_pt = st.text_input("Nama PT", "PT Contoh Sejahtera")
pejabat = st.text_input("Nama Pejabat", "Reza Fahlevi Lubis")

if coa_file and saldo_file and jurnal_file:
    coa = pd.read_excel(coa_file)
    saldo_awal = pd.read_excel(saldo_file)
    jurnal = pd.read_excel(jurnal_file)

    # Normalisasi nama kolom
    coa.columns = coa.columns.str.strip().str.lower()
    saldo_awal.columns = saldo_awal.columns.str.strip().str.lower()
    jurnal.columns = jurnal.columns.str.strip().str.lower()

    # Pastikan kolom saldo benar
    if "saldo_awal" not in saldo_awal.columns and "saldo" in saldo_awal.columns:
        saldo_awal = saldo_awal.rename(columns={"saldo": "saldo_awal"})

    saldo_awal["saldo_awal"] = pd.to_numeric(saldo_awal["saldo_awal"], errors="coerce").fillna(0)
    jurnal["debit"] = pd.to_numeric(jurnal["debit"], errors="coerce").fillna(0)
    jurnal["kredit"] = pd.to_numeric(jurnal["kredit"], errors="coerce").fillna(0)

    # Hitung saldo akhir
    saldo = saldo_awal.merge(coa, on="kode_akun", how="left")
    total_jurnal = jurnal.groupby("kode_akun")[["debit", "kredit"]].sum().reset_index()
    saldo = saldo.merge(total_jurnal, on="kode_akun", how="left").fillna(0)
    saldo["saldo_akhir"] = saldo.apply(lambda r: hitung_saldo(r["saldo_awal"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1)

    # Bagi berdasarkan laporan
    df_laba_rugi = {
        "Pendapatan": saldo[saldo["sub_tipe_laporan"]=="Pendapatan"],
        "Beban Umum Administrasi": saldo[saldo["sub_tipe_laporan"]=="Beban Umum Administrasi"],
        "Pendapatan Luar Usaha": saldo[saldo["sub_tipe_laporan"]=="Pendapatan Luar Usaha"],
        "Beban Luar Usaha": saldo[saldo["sub_tipe_laporan"]=="Beban Luar Usaha"]
    }

    laba_rugi = (df_laba_rugi["Pendapatan"]["saldo_akhir"].sum() +
                 df_laba_rugi["Pendapatan Luar Usaha"]["saldo_akhir"].sum() -
                 df_laba_rugi["Beban Umum Administrasi"]["saldo_akhir"].sum() -
                 df_laba_rugi["Beban Luar Usaha"]["saldo_akhir"].sum())

    df_aset = saldo[saldo["sub_tipe_laporan"].str.contains("Aset", na=False)]
    df_kewajiban = saldo[saldo["sub_tipe_laporan"].str.contains("Kewajiban", na=False)]
    df_ekuitas = saldo[saldo["sub_tipe_laporan"].str.contains("Ekuitas", na=False)]

    total_aset = df_aset["saldo_akhir"].sum()
    total_kewajiban = df_kewajiban["saldo_akhir"].sum()
    total_ekuitas = df_ekuitas["saldo_akhir"].sum()

    # Preview
    st.header("📑 Laporan Laba Rugi")
    for k, v in df_laba_rugi.items():
        st.subheader(k)
        st.write(v[["kode_akun","nama_akun","saldo_akhir"]])
        st.write(f"**TOTAL {k.upper()} : Rp {v['saldo_akhir'].sum():,.0f}**")
    st.success(f"LABA (RUGI) BERSIH : Rp {laba_rugi:,.0f}")

    st.header("📑 Laporan Posisi Keuangan (Neraca)")
    for title, df, total in [("ASET", df_aset, total_aset), ("KEWAJIBAN", df_kewajiban, total_kewajiban), ("EKUITAS", df_ekuitas, total_ekuitas)]:
        st.subheader(title)
        st.write(df[["kode_akun","nama_akun","saldo_akhir"]])
        st.write(f"**TOTAL {title.upper()} : Rp {total:,.0f}**")

    # ---------------- EXPORT ----------------
    st.subheader("📤 Export Laporan")
    # Excel
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        for sheet, df in df_laba_rugi.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
        df_aset.to_excel(writer, sheet_name="Aset", index=False)
        df_kewajiban.to_excel(writer, sheet_name="Kewajiban", index=False)
        df_ekuitas.to_excel(writer, sheet_name="Ekuitas", index=False)

    st.download_button("⬇️ Download Excel", data=output_excel.getvalue(), file_name="laporan_keuangan.xlsx")

    # PDF
    pdf_buffer = export_pdf(df_laba_rugi, laba_rugi, df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, pejabat)
    st.download_button("⬇️ Download PDF", data=pdf_buffer, file_name="laporan_keuangan.pdf")
