import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Aplikasi Akuntansi Otomatis", layout="wide")

st.title("📊 Aplikasi Akuntansi Otomatis")
st.caption("By: Reza Fahlevi Lubis BKP @zavibis")
st.markdown("""
Aplikasi ini membaca **COA**, **Saldo Awal**, dan **Jurnal Umum** dalam format Excel
untuk menghasilkan laporan keuangan otomatis: **Laba Rugi** dan **Neraca**.
Semua data diproses di memori — **tidak disimpan ke server.**
""")

st.divider()

# === UPLOAD FILE ===
coa_file   = st.file_uploader("📘 Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("💰 Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file= st.file_uploader("🧾 Upload Jurnal Umum.xlsx / Formimpor2.xlsx", type=["xlsx"])

if coa_file and saldo_file and jurnal_file:
    try:
        coa        = pd.read_excel(coa_file)
        saldo_awal = pd.read_excel(saldo_file)
        jurnal     = pd.read_excel(jurnal_file)

        st.success("✅ Semua file berhasil dibaca.")
        with st.expander("🔍 Preview Data COA"):
            st.dataframe(coa.head())
        with st.expander("🔍 Preview Saldo Awal"):
            st.dataframe(saldo_awal.head())
        with st.expander("🔍 Preview Jurnal Umum"):
            st.dataframe(jurnal.head())

        st.divider()

        # === NORMALISASI KOLOM ===
        coa.columns        = [c.strip() for c in coa.columns]
        saldo_awal.columns = [c.strip() for c in saldo_awal.columns]
        jurnal.columns     = [c.strip() for c in jurnal.columns]

        # --- Mapping fleksibel ---
        # COA
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

        # Saldo Awal
        for col in saldo_awal.columns:
            c = col.lower()
            if "kode" in c and "akun" in c: saldo_awal = saldo_awal.rename(columns={col:"kode_akun"})
            elif "saldo" in c: saldo_awal = saldo_awal.rename(columns={col:"saldo"})

        # Jurnal
        rename_jurnal = {}
        for col in jurnal.columns:
            c = col.lower()
            if "kode" in c and "akun" in c: rename_jurnal[col] = "kode_akun"
            elif "debit" in c: rename_jurnal[col] = "debit"
            elif "kredit" in c: rename_jurnal[col] = "kredit"
        jurnal = jurnal.rename(columns=rename_jurnal)

        # === VALIDASI DASAR ===
        if "kode_akun" not in coa.columns or "nama_akun" not in coa.columns:
            st.error("❌ Kolom 'Kode Akun' dan 'Nama Akun' wajib ada di COA.")
            st.stop()

        # === HITUNG SALDO AKHIR SESUAI NORMAL AKUN ===
        mutasi = jurnal.groupby("kode_akun")[["debit","kredit"]].sum().reset_index()
        df = coa.merge(saldo_awal, on="kode_akun", how="left").merge(mutasi, on="kode_akun", how="left")
        df[["saldo","debit","kredit"]] = df[["saldo","debit","kredit"]].fillna(0)

        # Rumus arah normal akun
        df["saldo_akhir"] = np.where(
            df["posisi_normal"].str.lower().str.strip()=="debit",
            df["saldo"] + df["debit"] - df["kredit"],
            df["saldo"] - df["debit"] + df["kredit"]
        )

        # Urut sesuai COA asli
        df["urutan"] = df.index

        # === PEMBAGIAN LAPORAN ===
        laporan_list = df["laporan"].dropna().unique().tolist()
        for jenis_laporan in laporan_list:
            st.header(f"📄 {jenis_laporan}")
            df_lap = df[df["laporan"]==jenis_laporan].copy()

            sub_groups = df_lap.groupby("sub_laporan", sort=False)
            grand_total = 0

            for nama_sub, group in sub_groups:
                st.markdown(f"### {nama_sub.upper()}")
                group = group.sort_values("urutan")

                # pisahkan header/detail
                detail = group[group["tipe_akun"].str.lower().str.contains("detail")]
                header = group[group["tipe_akun"].str.lower().str.contains("header")]

                total_sub = detail["saldo_akhir"].sum()
                grand_total += total_sub

                # tampilkan data detail
                if not detail.empty:
                    df_show = detail[["kode_akun","nama_akun","saldo_akhir"]].copy()
                    df_show["saldo_akhir"] = df_show["saldo_akhir"].map(lambda x: f"{x:,.0f}")
                    st.dataframe(df_show, hide_index=True, use_container_width=True)

                # tampilkan total sub tipe
                st.markdown(f"**TOTAL {nama_sub.upper()} : Rp {total_sub:,.0f}**")
                st.divider()

            st.subheader(f"💰 TOTAL {jenis_laporan.upper()} : Rp {grand_total:,.0f}")
            st.divider()

        # === EXPORT EXCEL ===
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Detail Akun")

        st.download_button(
            label="⬇️ Download Laporan (Excel)",
            data=output.getvalue(),
            file_name="Laporan_Keuangan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")

else:
    st.info("Unggah semua file untuk melanjutkan.")
