import streamlit as st
import pandas as pd

st.set_page_config(page_title="Laporan Keuangan", layout="wide")

st.title("📊 Laporan Keuangan")

# Upload file
coa_file = st.file_uploader("Upload COA.xlsx", type="xlsx")
saldo_file = st.file_uploader("Upload Saldo Awal.xlsx", type="xlsx")
jurnal_file = st.file_uploader("Upload Jurnal.xlsx", type="xlsx")

if coa_file and saldo_file and jurnal_file:
    # Load COA
    coa = pd.read_excel(coa_file)
    coa.columns = coa.columns.str.strip().str.lower()

    # Load Saldo Awal
    saldo_awal = pd.read_excel(saldo_file)
    saldo_awal.columns = saldo_awal.columns.str.strip().str.lower()
    saldo_awal = saldo_awal.rename(columns={"saldo": "saldo_awal"})

    # Load Jurnal
    jurnal = pd.read_excel(jurnal_file)
    jurnal.columns = jurnal.columns.str.strip().str.lower()

    # Pastikan numeric
    jurnal["debit"] = pd.to_numeric(jurnal["debit"], errors="coerce").fillna(0)
    jurnal["kredit"] = pd.to_numeric(jurnal["kredit"], errors="coerce").fillna(0)
    saldo_awal["saldo_awal"] = pd.to_numeric(saldo_awal["saldo_awal"], errors="coerce").fillna(0)

    # Aggregate debit & kredit per akun
    agg = jurnal.groupby("kode_akun")[["debit", "kredit"]].sum().reset_index()

    # Merge ke COA
    df = coa.merge(saldo_awal[["kode_akun", "saldo_awal"]], on="kode_akun", how="left")
    df = df.merge(agg, on="kode_akun", how="left").fillna({"saldo_awal": 0, "debit": 0, "kredit": 0})

    # Hitung saldo akhir
    def hitung_saldo(saldo_awal, debit, kredit, posisi_normal):
        if str(posisi_normal).lower() == "debit":
            return saldo_awal + debit - kredit
        else:
            return saldo_awal - debit + kredit

    df["saldo_akhir"] = df.apply(
        lambda r: hitung_saldo(r["saldo_awal"], r["debit"], r["kredit"], r["posisi_normal_akun"]),
        axis=1
    )

    # ===================== LAPORAN LABA RUGI =====================
    st.header("📑 Laporan Laba Rugi")

    # Filter pendapatan & beban
    pendapatan = df[(df["laporan"] == "Laporan Laba Rugi") & (df["sub_tipe_laporan"] == "Pendapatan")]
    beban_umum = df[(df["laporan"] == "Laporan Laba Rugi") & (df["sub_tipe_laporan"] == "Beban Umum Administrasi")]
    pendapatan_luar = df[(df["laporan"] == "Laporan Laba Rugi") & (df["sub_tipe_laporan"] == "Pendapatan Luar Usaha")]
    beban_luar = df[(df["laporan"] == "Laporan Laba Rugi") & (df["sub_tipe_laporan"] == "Beban Luar Usaha")]

    total_pendapatan = pendapatan["saldo_akhir"].sum()
    total_beban_umum = beban_umum["saldo_akhir"].sum()
    total_pendapatan_luar = pendapatan_luar["saldo_akhir"].sum()
    total_beban_luar = beban_luar["saldo_akhir"].sum()

    laba_rugi = total_pendapatan - total_beban_umum + total_pendapatan_luar - total_beban_luar

    st.subheader("Pendapatan")
    st.write(f"TOTAL PENDAPATAN : Rp {total_pendapatan:,.0f}")

    st.subheader("Beban Umum Administrasi")
    st.write(f"TOTAL BEBAN UMUM ADMINISTRASI : Rp {total_beban_umum:,.0f}")

    st.subheader("Pendapatan Luar Usaha")
    st.write(f"TOTAL PENDAPATAN LUAR USAHA : Rp {total_pendapatan_luar:,.0f}")

    st.subheader("Beban Luar Usaha")
    st.write(f"TOTAL BEBAN LUAR USAHA : Rp {total_beban_luar:,.0f}")

    st.success(f"💰 LABA (RUGI) BERSIH : Rp {laba_rugi:,.0f}")

    # ===================== LAPORAN POSISI KEUANGAN =====================
    st.header("📄 Laporan Posisi Keuangan (Neraca)")

    aset = df[(df["laporan"] == "Laporan Posisi Keuangan") & (df["sub_tipe_laporan"].str.contains("Aset", case=False, na=False))]
    kewajiban = df[(df["laporan"] == "Laporan Posisi Keuangan") & (df["sub_tipe_laporan"].str.contains("Kewajiban", case=False, na=False))]
    ekuitas = df[(df["laporan"] == "Laporan Posisi Keuangan") & (df["sub_tipe_laporan"].str.contains("Ekuitas", case=False, na=False))]

    total_aset = aset["saldo_akhir"].sum()
    total_kewajiban = kewajiban["saldo_akhir"].sum()
    total_ekuitas = ekuitas["saldo_akhir"].sum()

    st.subheader("ASET")
    st.dataframe(aset[["kode_akun", "nama_akun", "saldo_akhir"]])
    st.write(f"**TOTAL ASET : Rp {total_aset:,.0f}**")

    st.subheader("KEWAJIBAN")
    st.dataframe(kewajiban[["kode_akun", "nama_akun", "saldo_akhir"]])
    st.write(f"**TOTAL KEWAJIBAN : Rp {total_kewajiban:,.0f}**")

    st.subheader("EKUITAS")
    st.dataframe(ekuitas[["kode_akun", "nama_akun", "saldo_akhir"]])
    st.write(f"**TOTAL EKUITAS : Rp {total_ekuitas:,.0f}**")

    # Validasi Neraca
    st.info(f"TOTAL KEWAJIBAN + EKUITAS : Rp {total_kewajiban + total_ekuitas:,.0f}")
