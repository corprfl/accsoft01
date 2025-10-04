import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# ------------------------------
# Helper
# ------------------------------
def format_rupiah(x):
    try:
        return f"Rp {x:,.0f}".replace(",", ".")
    except:
        return "Rp 0"

def hitung_saldo(saldo_awal, debit, kredit, posisi):
    if str(posisi).lower().strip() == "debit":
        return saldo_awal + debit - kredit
    else:
        return saldo_awal - debit + kredit

# ------------------------------
# PDF Generator
# ------------------------------
def generate_pdf(df_sections, judul, periode, nama_pt, pejabat, jabatan):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{nama_pt}</b>", styles["Title"]))
    story.append(Paragraph(f"<b>{judul}</b>", styles["Heading2"]))
    story.append(Paragraph(f"Periode: {periode}", styles["Normal"]))
    story.append(Spacer(1, 12))

    for section, data in df_sections.items():
        story.append(Paragraph(f"<b>{section}</b>", styles["Heading3"]))
        if not data.empty:
            table_data = [["Akun", "Saldo"]]
            for _, row in data.iterrows():
                table_data.append([row["nama_akun"], format_rupiah(row["saldo_akhir"])])
            total = data["saldo_akhir"].sum()
            table_data.append([f"TOTAL {section.upper()}", format_rupiah(total)])

            t = Table(table_data, colWidths=[300, 150])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
                ("TEXTCOLOR",(0,0),(-1,0),colors.black),
                ("ALIGN",(1,1),(-1,-1),"RIGHT"),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0,0), (-1,0), 6),
                ("GRID",(0,0),(-1,-1),0.25,colors.black),
            ]))
            story.append(t)
        story.append(Spacer(1, 12))

    story.append(Spacer(1, 40))
    story.append(Paragraph(f"{jabatan},", styles["Normal"]))
    story.append(Spacer(1, 40))
    story.append(Paragraph(f"<u>{pejabat}</u>", styles["Normal"]))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# ------------------------------
# Streamlit App
# ------------------------------
st.title("📊 Aplikasi Akuntansi - Laporan Keuangan Profesional")

nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
periode = st.text_input("Periode", "31 Desember 2025")
pejabat = st.text_input("Nama Pejabat", "Reza Fahlevi Lubis")
jabatan = st.text_input("Jabatan", "Direktur")

uploaded_coa = st.file_uploader("Upload COA", type=["xlsx"])
uploaded_saldo = st.file_uploader("Upload Saldo Awal", type=["xlsx"])
uploaded_jurnal = st.file_uploader("Upload Jurnal Umum", type=["xlsx"])

if uploaded_coa and uploaded_saldo and uploaded_jurnal:
    coa = pd.read_excel(uploaded_coa)
    saldo_awal = pd.read_excel(uploaded_saldo)
    jurnal = pd.read_excel(uploaded_jurnal)

    # Normalisasi nama kolom
    coa.columns = coa.columns.str.strip().str.lower().str.replace(" ", "_")
    saldo_awal.columns = saldo_awal.columns.str.strip().str.lower().str.replace(" ", "_")
    jurnal.columns = jurnal.columns.str.strip().str.lower().str.replace(" ", "_")

    # Gabungkan data
    if "kode_akun" not in coa.columns:
        st.error("Kolom 'Kode Akun' tidak ditemukan di COA.xlsx")
        st.stop()

    df = coa.merge(saldo_awal, on="kode_akun", how="left").fillna(0)

    if "debit" not in jurnal.columns or "kredit" not in jurnal.columns:
        st.error("Kolom 'Debit' atau 'Kredit' tidak ditemukan di file jurnal.")
        st.stop()

    jurnal_group = jurnal.groupby("kode_akun")[["debit","kredit"]].sum().reset_index()
    df = df.merge(jurnal_group, on="kode_akun", how="left").fillna(0)

    # Deteksi kolom posisi normal (flexible)
    if "posisi_normal_akun" in df.columns:
        kol_posisi = "posisi_normal_akun"
    elif "posisi_normal" in df.columns:
        kol_posisi = "posisi_normal"
    else:
        st.error("Kolom 'Posisi Normal Akun' tidak ditemukan di COA.xlsx")
        st.stop()

    # Hitung saldo akhir aman
    df["saldo_akhir"] = df.apply(
        lambda r: hitung_saldo(r["saldo"], r["debit"], r["kredit"], r[kol_posisi]), axis=1
    )

    # Hitung laba rugi
    df_lr = df[df["laporan"].str.contains("Laba Rugi", case=False, na=False)].copy()
    total_pendapatan = df_lr[df_lr["sub_tipe_laporan"].str.contains("Pendapatan", case=False, na=False)]["saldo_akhir"].sum()
    total_beban = df_lr[df_lr["sub_tipe_laporan"].str.contains("Beban", case=False, na=False)]["saldo_akhir"].sum()
    laba_rugi = total_pendapatan - total_beban

    # Tambahkan ke akun 3004 (Laba Rugi Berjalan)
    if "3004" in df["kode_akun"].values:
        df.loc[df["kode_akun"]=="3004", "saldo_akhir"] += laba_rugi
    else:
        df = pd.concat([df, pd.DataFrame([{
            "kode_akun":"3004",
            "nama_akun":"Laba (Rugi) Berjalan",
            "tipe_akun":"Detail",
            kol_posisi:"Kredit",
            "laporan":"Laporan Posisi Keuangan",
            "sub_tipe_laporan":"Ekuitas",
            "saldo":0,"debit":0,"kredit":0,
            "saldo_akhir":laba_rugi
        }])], ignore_index=True)

    # Preview
    mode = st.radio("Tampilan Preview", ["Ringkas", "Detail"])

    if mode == "Detail":
        for laporan in df["laporan"].dropna().unique():
            st.subheader(laporan.upper())
            for sub in df[df["laporan"]==laporan]["sub_tipe_laporan"].dropna().unique():
                subset = df[(df["laporan"]==laporan) & (df["sub_tipe_laporan"]==sub)]
                st.markdown(f"### {sub}")
                st.dataframe(subset[["kode_akun","nama_akun","saldo_akhir"]])
                st.markdown(f"**TOTAL {sub.upper()} : {format_rupiah(subset['saldo_akhir'].sum())}**")
    else:
        for laporan in df["laporan"].dropna().unique():
            st.subheader(laporan.upper())
            ringkas = df[df["laporan"]==laporan].groupby("sub_tipe_laporan")["saldo_akhir"].sum().reset_index()
            ringkas.columns = ["Sub Laporan", "Total"]
            ringkas["Total"] = ringkas["Total"].apply(format_rupiah)
            st.dataframe(ringkas)

    # Export Excel
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Laporan", index=False)
    st.download_button(
        "⬇️ Download Excel", output_excel.getvalue(),
        "Laporan_Keuangan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="excel"
    )

    # Export PDF Laba Rugi
    sections_lr = {s: df[(df["laporan"].str.contains("Laba Rugi", case=False, na=False)) &
                         (df["sub_tipe_laporan"]==s)][["nama_akun","saldo_akhir"]] 
                   for s in df[df["laporan"].str.contains("Laba Rugi", case=False, na=False)]["sub_tipe_laporan"].dropna().unique()}
    pdf_lr = generate_pdf(sections_lr, "LAPORAN LABA RUGI", periode, nama_pt, pejabat, jabatan)
    st.download_button("⬇️ Download PDF Laba Rugi", pdf_lr, "Laporan_Laba_Rugi.pdf", mime="application/pdf", key="pdf_lr")

    # Export PDF Neraca
    sections_neraca = {s: df[(df["laporan"].str.contains("Posisi Keuangan", case=False, na=False)) &
                             (df["sub_tipe_laporan"]==s)][["nama_akun","saldo_akhir"]] 
                       for s in df[df["laporan"].str.contains("Posisi Keuangan", case=False, na=False)]["sub_tipe_laporan"].dropna().unique()}
    pdf_neraca = generate_pdf(sections_neraca, "LAPORAN POSISI KEUANGAN", periode, nama_pt, pejabat, jabatan)
    st.download_button("⬇️ Download PDF Neraca", pdf_neraca, "Laporan_Posisi_Keuangan.pdf", mime="application/pdf", key="pdf_neraca")

else:
    st.info("Unggah semua file (COA, Saldo Awal, dan Jurnal) untuk melanjutkan.")
