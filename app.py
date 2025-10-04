import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Aplikasi Akuntansi Otomatis", layout="wide")

st.title("📊 Aplikasi Akuntansi Otomatis")
st.caption("By: Reza Fahlevi Lubis BKP @zavibis")
st.markdown("""
Aplikasi ini membaca **COA**, **Saldo Awal**, dan **Jurnal Umum** dalam format Excel untuk menghasilkan laporan keuangan otomatis:  
**Laba Rugi, Neraca, dan Arus Kas** — tanpa menyimpan data ke server.
""")

st.divider()

# === UPLOAD FILE ===
coa_file = st.file_uploader("📘 Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("💰 Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file = st.file_uploader("🧾 Upload Jurnal Umum.xlsx / Formimpor2.xlsx", type=["xlsx"])

if coa_file and saldo_file and jurnal_file:
    try:
        coa = pd.read_excel(coa_file)
        saldo_awal = pd.read_excel(saldo_file)
        jurnal = pd.read_excel(jurnal_file)

        st.success("✅ Semua file berhasil dibaca.")
        with st.expander("🔍 Preview Data COA"):
            st.dataframe(coa.head())
        with st.expander("🔍 Preview Saldo Awal"):
            st.dataframe(saldo_awal.head())
        with st.expander("🔍 Preview Jurnal Umum"):
            st.dataframe(jurnal.head())

        st.divider()

        # === NORMALISASI KOLOM ===
        coa.columns = [c.lower().strip() for c in coa.columns]
        saldo_awal.columns = [c.lower().strip() for c in saldo_awal.columns]
        jurnal.columns = [c.lower().strip() for c in jurnal.columns]

        if 'kode_akun' not in coa.columns or 'nama_akun' not in coa.columns:
            st.error("COA harus memiliki kolom: 'kode_akun' dan 'nama_akun'.")
            st.stop()

        akun_df = coa[['kode_akun', 'nama_akun']].copy()

        # total mutasi
        mutasi = jurnal.groupby('kode_akun').agg({'debit': 'sum', 'kredit': 'sum'}).reset_index()

        df = akun_df.merge(saldo_awal, on='kode_akun', how='left').merge(mutasi, on='kode_akun', how='left')
        df = df.fillna(0)
        df['saldo_akhir'] = df['saldo'] + df['debit'] - df['kredit']

        # === KLASIFIKASI AKUN ===
        if 'tipe' not in coa.columns:
            coa['tipe'] = np.where(coa['kode_akun'].astype(str).str.startswith('4'), 'Pendapatan',
                        np.where(coa['kode_akun'].astype(str).str.startswith('5'), 'Beban',
                        np.where(coa['kode_akun'].astype(str).str.startswith('1'), 'Aset',
                        np.where(coa['kode_akun'].astype(str).str.startswith('2'), 'Kewajiban',
                        np.where(coa['kode_akun'].astype(str).str.startswith('3'), 'Ekuitas', 'Lainnya')))))

        df = df.merge(coa[['kode_akun','tipe']], on='kode_akun', how='left')

        # === LAPORAN LABA RUGI ===
        laba_rugi = df[df['tipe'].isin(['Pendapatan','Beban'])]
        total_pendapatan = laba_rugi[laba_rugi['tipe']=='Pendapatan']['saldo_akhir'].sum()
        total_beban = laba_rugi[laba_rugi['tipe']=='Beban']['saldo_akhir'].sum()
        laba_bersih = total_pendapatan - total_beban

        # === NERACA ===
        neraca = df[df['tipe'].isin(['Aset','Kewajiban','Ekuitas'])]

        # === ARUS KAS ===
        arus_kas = df[df['nama_akun'].str.contains('kas|bank', case=False, na=False)][['kode_akun','nama_akun','saldo_akhir']]

        st.header("📈 Laporan Laba Rugi")
        st.dataframe(laba_rugi[['kode_akun','nama_akun','saldo_akhir']])
        st.write(f"**Total Pendapatan:** Rp {total_pendapatan:,.0f}")
        st.write(f"**Total Beban:** Rp {total_beban:,.0f}")
        st.subheader(f"💰 Laba Bersih: Rp {laba_bersih:,.0f}")

        st.divider()
        st.header("📊 Neraca")
        st.dataframe(neraca[['kode_akun','nama_akun','saldo_akhir','tipe']])

        st.divider()
        st.header("💵 Arus Kas (Kas & Bank)")
        st.dataframe(arus_kas)

        # === EXPORT EXCEL ===
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            laba_rugi.to_excel(writer, index=False, sheet_name='Laba Rugi')
            neraca.to_excel(writer, index=False, sheet_name='Neraca')
            arus_kas.to_excel(writer, index=False, sheet_name='Arus Kas')
            df.to_excel(writer, index=False, sheet_name='Detail Akun')

        st.download_button(
            label="⬇️ Download Laporan Keuangan (Excel)",
            data=output.getvalue(),
            file_name="Laporan_Keuangan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")

else:
    st.info("Unggah semua file untuk melanjutkan.")
