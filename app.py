import streamlit as st
import pandas as pd
import io
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Aplikasi Akuntansi", layout="wide")

# ==============================
# Fungsi bantu
# ==============================
def bersihkan_kolom(df):
    df.columns = (
        df.columns
        .str.replace("\xa0", " ", regex=False)  
        .str.replace(r"[^0-9a-zA-Z_ ]", "", regex=True)  
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )
    return df

def hitung_saldo(saldo_awal, debit, kredit, posisi):
    if posisi and str(posisi).lower().startswith("debit"):
        return saldo_awal + debit - kredit
    elif posisi and str(posisi).lower().startswith("kredit"):
        return saldo_awal - debit + kredit
    else:
        return saldo_awal + debit - kredit

def format_rupiah(x):
    return f"Rp {x:,.0f}".replace(",", ".")

# ==============================
# Upload file
# ==============================
st.title("📊 Aplikasi Akuntansi Streamlit")

uploaded_coa = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
uploaded_saldo = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
uploaded_jurnal = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])

if uploaded_coa and uploaded_saldo and uploaded_jurnal:
    coa = bersihkan_kolom(pd.read_excel(uploaded_coa))
    saldo_awal = bersihkan_kolom(pd.read_excel(uploaded_saldo))
    jurnal = bersihkan_kolom(pd.read_excel(uploaded_jurnal))

    # pastikan kode akun jadi string
    for df in [coa, saldo_awal, jurnal]:
        if "kode_akun" in df.columns:
            df["kode_akun"] = df["kode_akun"].astype(str).str.strip()

    # normalisasi kolom saldo awal
    if "saldo" not in saldo_awal.columns:
        for col in saldo_awal.columns:
            if "saldo" in col.lower():
                saldo_awal.rename(columns={col: "saldo"}, inplace=True)

    # normalisasi debit kredit
    if "debit" not in jurnal.columns:
        for col in jurnal.columns:
            if "debit" in col.lower():
                jurnal.rename(columns={col: "debit"}, inplace=True)
    if "kredit" not in jurnal.columns:
        for col in jurnal.columns:
            if "kredit" in col.lower():
                jurnal.rename(columns={col: "kredit"}, inplace=True)

    # preview debug
    st.subheader("📋 Debug Preview Data")
    st.write("COA:", coa.head())
    st.write("Saldo Awal:", saldo_awal.head())
    st.write("Jurnal:", jurnal.head())

    # agregasi jurnal
    jurnal_agg = jurnal.groupby("kode_akun").agg({"debit":"sum", "kredit":"sum"}).reset_index()

    # merge
    df = coa.copy()
    if "kode_akun" in saldo_awal.columns and "saldo" in saldo_awal.columns:
        df = df.merge(saldo_awal[["kode_akun","saldo"]], on="kode_akun", how="left")
    else:
        df["saldo"] = 0
    df = df.merge(jurnal_agg, on="kode_akun", how="left").fillna(0)

    # hitung saldo akhir
    df["saldo_akhir"] = df.apply(
        lambda r: hitung_saldo(r["saldo"], r["debit"], r["kredit"], r.get("posisi_normal_akun","")),
        axis=1
    )

    # ==========================
    # LAPORAN LABA RUGI
    # ==========================
    st.header("📈 Laporan Laba Rugi")

    df_lr = df[df["laporan"]=="Laba Rugi"].copy()
    laba_rugi = df_lr["saldo_akhir"].sum()

    for header in df_lr["sub_tipe_laporan"].unique():
        sub = df_lr[df_lr["sub_tipe_laporan"]==header]
        total = sub["saldo_akhir"].sum()
        st.subheader(header.upper())
        st.table(sub[["kode_akun","nama_akun","saldo_akhir"]])
        st.write(f"**TOTAL {header.upper()} : {format_rupiah(total)}**")

    st.markdown(f"### 💰 LABA (RUGI) BERSIH : {format_rupiah(laba_rugi)}")

    # ==========================
    # LAPORAN POSISI KEUANGAN
    # ==========================
    st.header("📑 Laporan Posisi Keuangan (Neraca)")

    df_nr = df[df["laporan"].isin(["Aset","Kewajiban","Ekuitas"])].copy()

    # update akun 3004 dengan laba rugi
    if "3004" in df_nr["kode_akun"].astype(str).values:
        df_nr.loc[df_nr["kode_akun"].astype(str)=="3004","saldo_akhir"] = laba_rugi

    total_aset = df_nr[df_nr["laporan"]=="Aset"]["saldo_akhir"].sum()
    total_kewajiban = df_nr[df_nr["laporan"]=="Kewajiban"]["saldo_akhir"].sum()
    total_ekuitas = df_nr[df_nr["laporan"]=="Ekuitas"]["saldo_akhir"].sum()

    for header in ["Aset","Kewajiban","Ekuitas"]:
        sub = df_nr[df_nr["laporan"]==header]
        total = sub["saldo_akhir"].sum()
        st.subheader(header.upper())
        st.table(sub[["kode_akun","nama_akun","saldo_akhir"]])
        st.write(f"**TOTAL {header.upper()} : {format_rupiah(total)}**")

    st.markdown(f"### ✅ TOTAL ASET : {format_rupiah(total_aset)}")
    st.markdown(f"### ✅ TOTAL KEWAJIBAN + EKUITAS : {format_rupiah(total_kewajiban+total_ekuitas)}")
