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
        total_pendapatan = df_lr[df_lr["sub_laporan"].str.contains("pendapatan", case=False, na=False)]["saldo_akhir"].sum()
        total_beban = df_lr[df_lr["sub_laporan"].str.contains("beban", case=False, na=False)]["saldo_akhir"].sum()
        laba_rugi = total_pendapatan - abs(total_beban)

        # === TAMPILKAN LABA RUGI ===
        st.header(f"🏦 LAPORAN LABA RUGI - {periode}")
        for sub, group in df_lr.groupby("sub_laporan"):
            detail = group[group["tipe_akun"].str.lower().str.contains("detail")]
            subtotal = detail["saldo_akhir"].sum()
            if preview_mode == "Detail":
                st.dataframe(detail[["kode_akun","nama_akun","saldo_akhir"]])
            st.markdown(f"**TOTAL {sub.upper()} : Rp {fmt_rupiah(subtotal)}**")
        st.subheader(f"💰 LABA (RUGI) BERSIH : Rp {fmt_rupiah(laba_rugi)}")

        # === TAMPILKAN NERACA ===
        st.header(f"📒 LAPORAN POSISI KEUANGAN - {periode}")
        df_neraca = df[df["laporan"].str.contains("Posisi Keuangan", case=False, na=False)].copy()

        # Tambahkan laba rugi ke akun 3004 (Laba Rugi Berjalan)
        if "3004" in df_neraca["kode_akun"].values:
            df_neraca.loc[df_neraca["kode_akun"]=="3004", "saldo_akhir"] += laba_rugi
        elif df_neraca["nama_akun"].str.contains("laba", case=False).any():
            idx = df_neraca[df_neraca["nama_akun"].str.contains("laba", case=False)].index[0]
            df_neraca.loc[idx, "saldo_akhir"] += laba_rugi
        else:
            # fallback buat bikin akun baru
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

        for sub, group in df_neraca.groupby("sub_laporan"):
            detail = group[group["tipe_akun"].str.lower().str.contains("detail")]
            subtotal = detail["saldo_akhir"].sum()
            if preview_mode == "Detail":
                st.dataframe(detail[["kode_akun","nama_akun","saldo_akhir"]])
            st.markdown(f"**TOTAL {sub.upper()} : Rp {fmt_rupiah(subtotal)}**")

        total_aset = df_neraca[df_neraca["sub_laporan"].str.contains("aset", case=False, na=False)]["saldo_akhir"].sum()
        total_liab_ekuitas = df_neraca[df_neraca["sub_laporan"].str.contains("kewajiban|ekuitas", case=False, na=False)]["saldo_akhir"].sum()
        st.subheader(f"TOTAL ASET : Rp {fmt_rupiah(total_aset)}")
        st.subheader(f"TOTAL KEWAJIBAN + EKUITAS : Rp {fmt_rupiah(total_liab_ekuitas)}")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")

else:
    st.info("Unggah semua file untuk melanjutkan.")
