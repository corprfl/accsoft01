import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Aplikasi Akuntansi Otomatis", layout="wide")

st.title("📊 Aplikasi Akuntansi Otomatis")
st.caption("By: Reza Fahlevi Lubis BKP @zavibis")

# === Input custom header laporan ===
with st.sidebar:
    st.header("⚙️ Pengaturan Laporan")
    nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
    periode = st.text_input("Periode Laporan", "31 Desember 2025")
    pejabat = st.text_input("Nama Pejabat", "Reza Fahlevi Lubis")
    jabatan = st.text_input("Jabatan Pejabat", "Direktur")
    preview_mode = st.radio("Preview Mode", ["Total", "Detail"])

# === UPLOAD FILE ===
coa_file   = st.file_uploader("📘 Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("💰 Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file= st.file_uploader("🧾 Upload Jurnal Umum.xlsx / Formimpor2.xlsx", type=["xlsx"])

def fmt_rupiah(val):
    if val < 0:
        return f"({abs(val):,.0f})"
    return f"{val:,.0f}"

def build_pdf(jenis, df, subtotal_dict, total_label, total_value, nama_pt, periode, pejabat, jabatan):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # Header
    elements.append(Paragraph(f"<b>{nama_pt}</b>", styles['Title']))
    elements.append(Paragraph(f"{jenis}", styles['Heading2']))
    elements.append(Paragraph(f"Periode: {periode}", styles['Normal']))
    elements.append(Spacer(1, 12))

    # Detail
    for sub, subtotal in subtotal_dict.items():
        elements.append(Paragraph(f"<b>{sub.upper()}</b>", styles['Normal']))
        sub_df = df[df["sub_laporan"]==sub]
        data = [["Akun", "Saldo"]]
        for _, row in sub_df.iterrows():
            if row["tipe_akun"].lower() == "detail":
                data.append([row["nama_akun"], fmt_rupiah(row["saldo_akhir"])])
        data.append([f"TOTAL {sub.upper()}", fmt_rupiah(subtotal)])
        t = Table(data, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.grey),
            ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
            ("ALIGN",(1,1),(-1,-1),"RIGHT"),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("BOTTOMPADDING",(0,0),(-1,0),6),
            ("GRID",(0,0),(-1,-1),0.25,colors.black)
        ]))
        elements.append(t)
        elements.append(Spacer(1, 12))

    # Total besar
    elements.append(Paragraph(f"<b>{total_label} : Rp {fmt_rupiah(total_value)}</b>", styles['Heading2']))
    elements.append(Spacer(1, 50))
    elements.append(Paragraph(f"{jabatan},", styles['Normal']))
    elements.append(Spacer(1, 30))
    elements.append(Paragraph(f"<u>{pejabat}</u>", styles['Normal']))

    doc.build(elements)
    return buffer.getvalue()

if coa_file and saldo_file and jurnal_file:
    try:
        coa        = pd.read_excel(coa_file)
        saldo_awal = pd.read_excel(saldo_file)
        jurnal     = pd.read_excel(jurnal_file)

        # --- Normalisasi kolom ---
        coa.columns        = [c.strip() for c in coa.columns]
        saldo_awal.columns = [c.strip() for c in saldo_awal.columns]
        jurnal.columns     = [c.strip() for c in jurnal.columns]

        # Mapping COA
        rename_coa = {}
        for col in coa.columns:
            c = col.lower()
            if "kode" in c and "akun" in c: rename_coa[col] = "kode_akun"
            elif "nama" in c and "akun" in c: rename_coa[col] = "nama_akun"
            elif "tipe" in c and "akun" in c: rename_coa[col] = "tipe_akun"
            elif "normal" in c: rename_coa[col] = "posisi_normal"
            elif "laporan" == c: rename_coa[col] = "laporan"
            elif "sub" in c and "laporan" in c: rename_coa[col] = "sub_laporan"
        coa = coa.rename(columns=rename_coa)

        # Mapping Saldo Awal
        for col in saldo_awal.columns:
            c = col.lower()
            if "kode" in c and "akun" in c: saldo_awal = saldo_awal.rename(columns={col:"kode_akun"})
            elif "saldo" in c: saldo_awal = saldo_awal.rename(columns={col:"saldo"})

        # Mapping Jurnal
        rename_jurnal = {}
        for col in jurnal.columns:
            c = col.lower()
            if "kode" in c and "akun" in c: rename_jurnal[col] = "kode_akun"
            elif "debit" in c: rename_jurnal[col] = "debit"
            elif "kredit" in c: rename_jurnal[col] = "kredit"
        jurnal = jurnal.rename(columns=rename_jurnal)

        # --- Hitung saldo akhir ---
        mutasi = jurnal.groupby("kode_akun")[["debit","kredit"]].sum().reset_index()
        df = coa.merge(saldo_awal, on="kode_akun", how="left").merge(mutasi, on="kode_akun", how="left")
        df[["saldo","debit","kredit"]] = df[["saldo","debit","kredit"]].fillna(0)

        df["saldo_akhir"] = np.where(
            df["posisi_normal"].str.lower().str.strip()=="debit",
            df["saldo"] + df["debit"] - df["kredit"],
            df["saldo"] - df["debit"] + df["kredit"]
        )

        # ✅ Akumulasi Penyusutan di Neraca → negatif
        mask_akum = (df["laporan"].str.contains("Posisi Keuangan", case=False, na=False)) & \
                    (df["nama_akun"].str.contains("akum", case=False, na=False))
        df.loc[mask_akum, "saldo_akhir"] *= -1

        # ✅ Prive → negatif (pengurang ekuitas)
        mask_prive = (df["laporan"].str.contains("Posisi Keuangan", case=False, na=False)) & \
                     (df["nama_akun"].str.contains("prive", case=False, na=False))
        df.loc[mask_prive, "saldo_akhir"] *= -1

        # --- Hitung Laba Rugi Bersih ---
        df_lr = df[df["laporan"].str.contains("Laba Rugi", case=False, na=False)]
        subtotal_lr = {}
        for sub, group in df_lr.groupby("sub_laporan"):
            subtotal_lr[sub] = group[group["tipe_akun"].str.lower().str.contains("detail")]["saldo_akhir"].sum()
        total_pendapatan = sum(v for k,v in subtotal_lr.items() if "pendapatan" in k.lower())
        total_beban = sum(v for k,v in subtotal_lr.items() if "beban" in k.lower())
        laba_rugi = total_pendapatan - abs(total_beban)

        # === Preview Laba Rugi ===
        st.header(f"🏦 LAPORAN LABA RUGI - {periode}")
        for sub, subtotal in subtotal_lr.items():
            if preview_mode == "Detail":
                st.dataframe(df_lr[df_lr["sub_laporan"]==sub][["kode_akun","nama_akun","saldo_akhir"]])
            st.markdown(f"**TOTAL {sub.upper()} : Rp {fmt_rupiah(subtotal)}**")
        st.subheader(f"💰 LABA (RUGI) BERSIH : Rp {fmt_rupiah(laba_rugi)}")

        # === Neraca ===
        st.header(f"📒 LAPORAN POSISI KEUANGAN - {periode}")
        df_neraca = df[df["laporan"].str.contains("Posisi Keuangan", case=False, na=False)].copy()

        # Tambahkan laba rugi ke akun 3004 (Laba Rugi Berjalan)
        if "3004" in df_neraca["kode_akun"].values:
            df_neraca.loc[df_neraca["kode_akun"]=="3004", "saldo_akhir"] += laba_rugi
        else:
            df_neraca = pd.concat([df_neraca, pd.DataFrame([{
                "kode_akun":"3004",
                "nama_akun":"Laba (Rugi) Berjalan",
                "tipe_akun":"Detail",
                "posisi_normal":"Kredit",
                "laporan":"Laporan Posisi Keuangan",
                "sub_laporan":"Ekuitas",
                "saldo":0,"debit":0,"kredit":0,
                "saldo_akhir":laba_rugi
            }])], ignore_index=True)

        subtotal_neraca = {}
        for sub, group in df_neraca.groupby("sub_laporan"):
            subtotal_neraca[sub] = group[group["tipe_akun"].str.lower().str.contains("detail")]["saldo_akhir"].sum()

        total_aset = sum(v for k,v in subtotal_neraca.items() if "aset" in k.lower())
        total_liab_ekuitas = sum(v for k,v in subtotal_neraca.items() if "kewajiban" in k.lower() or "ekuitas" in k.lower())

        for sub, subtotal in subtotal_neraca.items():
            if preview_mode == "Detail":
                st.dataframe(df_neraca[df_neraca["sub_laporan"]==sub][["kode_akun","nama_akun","saldo_akhir"]])
            st.markdown(f"**TOTAL {sub.upper()} : Rp {fmt_rupiah(subtotal)}**")

        st.subheader(f"TOTAL ASET : Rp {fmt_rupiah(total_aset)}")
        st.subheader(f"TOTAL KEWAJIBAN + EKUITAS : Rp {fmt_rupiah(total_liab_ekuitas)}")

        # === Export Excel ===
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            df_lr.to_excel(writer, index=False, sheet_name="Laba Rugi")
            df_neraca.to_excel(writer, index=False, sheet_name="Neraca")
        st.download_button("⬇️ Download Excel", output_excel.getvalue(), "Laporan_Keuangan.xlsx")

        # === Export PDF terpisah ===
        pdf_lr = build_pdf("LAPORAN LABA RUGI", df_lr, subtotal_lr, "LABA (RUGI) BERSIH", laba_rugi, nama_pt, periode, pejabat, jabatan)
        pdf_neraca = build_pdf("LAPORAN POSISI KEUANGAN", df_neraca, subtotal_neraca, "TOTAL ASET = TOTAL KEWAJIBAN + EKUITAS", total_aset, nama_pt, periode, pejabat, jabatan)

        st.download_button("⬇️ Download PDF Laba Rugi", pdf_lr, "Laporan_Laba_Rugi.pdf")
        st.download_button("⬇️ Download PDF Neraca", pdf_neraca, "Laporan_Posisi_Keuangan.pdf")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

else:
    st.info("Unggah semua file untuk melanjutkan.")
