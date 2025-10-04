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
    if posisi.lower().startswith("debit"):
        return saldo_awal + debit - kredit
    elif posisi.lower().startswith("kredit"):
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

    # pastikan ada kolom
    if "kode_akun" not in coa.columns or "posisi_normal_akun" not in coa.columns:
        st.error("COA harus punya kolom: kode_akun dan posisi_normal_akun")
        st.stop()

    # agregasi jurnal
    jurnal_agg = jurnal.groupby("kode_akun").agg({"debit":"sum", "kredit":"sum"}).reset_index()

    # merge ke master coa
    df = coa.copy()
    if "kode_akun" in saldo_awal.columns and "saldo" in saldo_awal.columns:
        df = df.merge(saldo_awal[["kode_akun","saldo"]], on="kode_akun", how="left")
    else:
        df["saldo"] = 0
    df = df.merge(jurnal_agg, on="kode_akun", how="left").fillna(0)

    # hitung saldo akhir
    df["saldo_akhir"] = df.apply(
        lambda r: hitung_saldo(r["saldo"], r["debit"], r["kredit"], r["posisi_normal_akun"]),
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

    # ==========================
    # EXPORT PDF & EXCEL
    # ==========================
    st.subheader("⬇️ Export Laporan")

    # Export Excel
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        df_lr.to_excel(writer, sheet_name="Laba Rugi", index=False)
        df_nr.to_excel(writer, sheet_name="Neraca", index=False)
    st.download_button(
        "📥 Download Excel",
        data=output_excel.getvalue(),
        file_name="laporan_keuangan.xlsx"
    )

    # Export PDF
    output_pdf = io.BytesIO()
    c = canvas.Canvas(output_pdf, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(300, 800, "PT Contoh Sejahtera")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 770, "LAPORAN LABA RUGI")
    c.drawString(50, 750, f"Laba (Rugi) Bersih: {format_rupiah(laba_rugi)}")
    c.drawString(50, 730, "LAPORAN POSISI KEUANGAN")
    c.drawString(50, 710, f"Total Aset: {format_rupiah(total_aset)}")
    c.drawString(50, 690, f"Total Kewajiban + Ekuitas: {format_rupiah(total_kewajiban+total_ekuitas)}")
    c.showPage()
    c.save()

    st.download_button(
        "📄 Download PDF",
        data=output_pdf.getvalue(),
        file_name="laporan_keuangan.pdf",
        mime="application/pdf"
    )
