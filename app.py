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

        # Akumulasi Penyusutan jadi negatif
        df.loc[df["nama_akun"].str.contains("akum|peny", case=False, na=False), "saldo_akhir"] *= -1

        # Urut sesuai COA
        df["urutan"] = df.index

        # === TAMPILKAN LAPORAN ===
        laporan_list = df["laporan"].dropna().unique().tolist()
        for jenis_laporan in laporan_list:
            st.header(f"📄 {jenis_laporan} - {periode}")
            df_lap = df[df["laporan"]==jenis_laporan].copy()

            sub_groups = df_lap.groupby("sub_laporan", sort=False)

            total_laporan = 0
            for nama_sub, group in sub_groups:
                st.markdown(f"### {nama_sub.upper()}")
                detail = group[group["tipe_akun"].str.lower().str.contains("detail")]
                subtotal = detail["saldo_akhir"].sum()
                total_laporan += subtotal

                if preview_mode == "Detail":
                    st.dataframe(detail[["kode_akun","nama_akun","saldo_akhir"]])
                st.markdown(f"**TOTAL {nama_sub.upper()} : Rp {subtotal:,.0f}**")
                st.divider()

            if "laba rugi" in jenis_laporan.lower():
                total_pendapatan = df_lap[df_lap["sub_laporan"].str.contains("pendapatan", case=False, na=False)]["saldo_akhir"].sum()
                total_beban = df_lap[df_lap["sub_laporan"].str.contains("beban", case=False, na=False)]["saldo_akhir"].sum()
                laba_rugi = total_pendapatan - abs(total_beban)
                st.subheader(f"💰 LABA (RUGI) BERSIH : Rp {laba_rugi:,.0f}")
            else:
                st.subheader(f"💰 TOTAL {jenis_laporan.upper()} : Rp {total_laporan:,.0f}")

        # === EXPORT EXCEL ===
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            for jenis_laporan in laporan_list:
                df[df["laporan"]==jenis_laporan].to_excel(writer, index=False, sheet_name=jenis_laporan[:30])
        st.download_button(
            label="⬇️ Download Laporan Excel",
            data=output_excel.getvalue(),
            file_name="Laporan_Keuangan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # === EXPORT PDF ===
        output_pdf = BytesIO()
        doc = SimpleDocTemplate(output_pdf, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph(f"<b>{nama_pt}</b>", styles['Title']))
        elements.append(Paragraph(f"{jenis_laporan}", styles['Heading2']))
        elements.append(Paragraph(f"Periode: {periode}", styles['Normal']))
        elements.append(Spacer(1,12))

        for jenis_laporan in laporan_list:
            elements.append(Paragraph(f"<b>{jenis_laporan}</b>", styles['Heading2']))
            df_lap = df[df["laporan"]==jenis_laporan].copy()
            sub_groups = df_lap.groupby("sub_laporan", sort=False)

            for nama_sub, group in sub_groups:
                elements.append(Paragraph(f"<b>{nama_sub.upper()}</b>", styles['Normal']))
                detail = group[group["tipe_akun"].str.lower().str.contains("detail")]
                subtotal = detail["saldo_akhir"].sum()

                data = [["Akun","Saldo"]]
                for _, row in detail.iterrows():
                    data.append([row["nama_akun"], f"{row['saldo_akhir']:,.0f}"])
                data.append([f"TOTAL {nama_sub.upper()}", f"{subtotal:,.0f}"])

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
                elements.append(Spacer(1,12))

        elements.append(Spacer(1,50))
        elements.append(Paragraph(f"{jabatan},", styles['Normal']))
        elements.append(Spacer(1,30))
        elements.append(Paragraph(f"<u>{pejabat}</u>", styles['Normal']))

        doc.build(elements)
        st.download_button(
            label="⬇️ Download Laporan PDF",
            data=output_pdf.getvalue(),
            file_name="Laporan_Keuangan.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")

else:
    st.info("Unggah semua file untuk melanjutkan.")
