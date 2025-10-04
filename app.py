import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from datetime import datetime

st.set_page_config(page_title="Aplikasi Laporan Keuangan", layout="wide")

# ==================================
# Fungsi bantu
# ==================================
def bersihkan_kolom(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"[^0-9a-zA-Z_ ]", "", regex=True)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )
    return df

def hitung_saldo(saldo_awal, debit, kredit, posisi_normal):
    if str(posisi_normal).lower().startswith("debit"):
        return saldo_awal + debit - kredit
    elif str(posisi_normal).lower().startswith("kredit"):
        return saldo_awal - debit + kredit
    else:
        return saldo_awal + debit - kredit

def format_rupiah(x):
    return f"Rp {x:,.0f}".replace(",", ".")

# ==================================
# Export PDF Laba Rugi
# ==================================
def export_pdf_laba_rugi(df_lr, laba_rugi, nama_pt, pejabat, tanggal_akhir):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    y = h - 2.5*cm

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 0.6*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, y, "LAPORAN LABA RUGI")
    y -= 0.5*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, f"Untuk Periode yang Berakhir Pada {tanggal_akhir}")
    y -= 1*cm

    for judul, df in df_lr.items():
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, judul)
        y -= 0.35*cm
        c.setFont("Helvetica", 9)
        for _, r in df.iterrows():
            c.drawString(2.3*cm, y, str(r["nama_akun"]))
            c.drawRightString(w-2*cm, y, f"Rp {r['saldo_akhir']:,.0f}")
            y -= 0.3*cm
        c.setFont("Helvetica-Bold", 9)
        c.drawString(2*cm, y, f"TOTAL {judul.upper()}")
        c.drawRightString(w-2*cm, y, f"Rp {df['saldo_akhir'].sum():,.0f}")
        y -= 0.5*cm

    c.setFont("Helvetica-Bold", 11)
    c.drawString(2*cm, y, f"LABA (RUGI) BERSIH : {format_rupiah(laba_rugi)}")
    c.setFont("Helvetica", 10)
    c.drawString(w-8*cm, 2.5*cm, "Direktur,")
    c.drawString(w-8*cm, 1.8*cm, pejabat)
    c.save()
    buf.seek(0)
    return buf

# ==================================
# Export PDF Neraca
# ==================================
def export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas,
                      total_aset, total_kewajiban, total_ekuitas,
                      nama_pt, pejabat, tanggal_akhir):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    y = h - 2.5*cm

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 0.6*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, y, "LAPORAN POSISI KEUANGAN (NERACA)")
    y -= 0.5*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, f"Per {tanggal_akhir}")
    y -= 1*cm

    def draw_sec(title, df, total):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, title)
        y -= 0.35*cm
        c.setFont("Helvetica", 9)
        saldo_col = "saldo_akhir_adj" if "saldo_akhir_adj" in df.columns else "saldo_akhir"
        for _, r in df.iterrows():
            c.drawString(2.3*cm, y, str(r["nama_akun"]))
            c.drawRightString(w-2*cm, y, f"Rp {r[saldo_col]:,.0f}")
            y -= 0.3*cm
        c.setFont("Helvetica-Bold", 9)
        c.drawString(2*cm, y, f"TOTAL {title.upper()}")
        c.drawRightString(w-2*cm, y, f"Rp {total:,.0f}")
        y -= 0.6*cm

    draw_sec("ASET", df_aset, total_aset)
    draw_sec("KEWAJIBAN", df_kewajiban, total_kewajiban)
    draw_sec("EKUITAS", df_ekuitas, total_ekuitas)

    c.setFont("Helvetica", 10)
    c.drawString(w-8*cm, 2.5*cm, "Direktur,")
    c.drawString(w-8*cm, 1.8*cm, pejabat)
    c.save()
    buf.seek(0)
    return buf

# ==================================
# APP STREAMLIT
# ==================================
st.title("📊 Generator Laporan Keuangan Profesional")

coa_file = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])

nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
pejabat = st.text_input("Nama Pejabat TTD", "Reza Fahlevi Lubis")
col1, col2 = st.columns(2)
with col1:
    tanggal_awal = st.date_input("Tanggal Awal Periode", datetime(2025,1,1))
with col2:
    tanggal_akhir = st.date_input("Tanggal Akhir Periode", datetime(2025,12,31))
periode_text = tanggal_akhir.strftime("%d %B %Y")

if coa_file and saldo_file and jurnal_file:
    coa = bersihkan_kolom(pd.read_excel(coa_file))
    saldo_awal = bersihkan_kolom(pd.read_excel(saldo_file))
    jurnal = bersihkan_kolom(pd.read_excel(jurnal_file))

    if "saldo_awal" not in saldo_awal.columns:
        candidates = [c for c in saldo_awal.columns if "saldo" in c]
        if candidates:
            saldo_awal.rename(columns={candidates[0]: "saldo_awal"}, inplace=True)
        else:
            st.error(f"Tidak ada kolom saldo. Kolom tersedia: {list(saldo_awal.columns)}")
            st.stop()

    saldo_awal["saldo_awal"] = pd.to_numeric(saldo_awal["saldo_awal"], errors="coerce").fillna(0)
    jurnal["debit"] = pd.to_numeric(jurnal.get("debit", 0), errors="coerce").fillna(0)
    jurnal["kredit"] = pd.to_numeric(jurnal.get("kredit", 0), errors="coerce").fillna(0)

    jurnal_sum = jurnal.groupby("kode_akun")[["debit","kredit"]].sum().reset_index()
    df = coa.merge(saldo_awal[["kode_akun","saldo_awal"]], on="kode_akun", how="left").fillna(0)
    df = df.merge(jurnal_sum, on="kode_akun", how="left").fillna(0)

    df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(r["saldo_awal"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1)

    df_lr = {
        "Pendapatan": df[df["sub_tipe_laporan"]=="Pendapatan"],
        "Beban Umum Administrasi": df[df["sub_tipe_laporan"]=="Beban Umum Administrasi"],
        "Pendapatan Luar Usaha": df[df["sub_tipe_laporan"]=="Pendapatan Luar Usaha"],
        "Beban Luar Usaha": df[df["sub_tipe_laporan"]=="Beban Luar Usaha"]
    }

    laba_rugi = (
        df_lr["Pendapatan"]["saldo_akhir"].sum()
        + df_lr["Pendapatan Luar Usaha"]["saldo_akhir"].sum()
        - df_lr["Beban Umum Administrasi"]["saldo_akhir"].sum()
        - df_lr["Beban Luar Usaha"]["saldo_akhir"].sum()
    )

    df_aset = df[df["sub_tipe_laporan"].str.contains("Aset", na=False)]
    df_kewajiban = df[df["sub_tipe_laporan"].str.contains("Kewajiban", na=False)]
    df_ekuitas = df[df["sub_tipe_laporan"].str.contains("Ekuitas", na=False)].copy()

    df_ekuitas["saldo_akhir_adj"] = df_ekuitas.apply(
        lambda r: r["saldo_akhir"] if str(r["posisi_normal_akun"]).lower()=="kredit" else -r["saldo_akhir"],
        axis=1
    )

    if "3004" in df_ekuitas["kode_akun"].astype(str).values:
        df_ekuitas.loc[df_ekuitas["kode_akun"].astype(str)=="3004","saldo_akhir_adj"] = laba_rugi

    total_aset = df_aset["saldo_akhir"].sum()
    total_kewajiban = df_kewajiban["saldo_akhir"].sum()
    total_ekuitas = df_ekuitas["saldo_akhir_adj"].sum()

    # PREVIEW
    mode = st.radio("Mode Preview", ["Total", "Detail"], horizontal=True)
    st.header("📈 Laporan Laba Rugi")
    for judul, data in df_lr.items():
        st.subheader(judul)
        if mode=="Detail": st.dataframe(data[["kode_akun","nama_akun","saldo_akhir"]])
        st.write(f"**TOTAL {judul.upper()} : {format_rupiah(data['saldo_akhir'].sum())}**")
    st.success(f"💰 LABA (RUGI) BERSIH : {format_rupiah(laba_rugi)}")

    st.header("📊 Neraca (Posisi Keuangan)")
    for title, data, total in [("ASET", df_aset, total_aset),
                               ("KEWAJIBAN", df_kewajiban, total_kewajiban),
                               ("EKUITAS", df_ekuitas, total_ekuitas)]:
        st.subheader(title)
        col = "saldo_akhir_adj" if "saldo_akhir_adj" in data.columns else "saldo_akhir"
        if mode=="Detail": st.dataframe(data[["kode_akun","nama_akun",col]])
        st.write(f"**TOTAL {title.upper()} : {format_rupiah(total)}**")

    st.info(f"TOTAL ASET : {format_rupiah(total_aset)} | TOTAL KEWAJIBAN + EKUITAS : {format_rupiah(total_kewajiban + total_ekuitas)}")

    # EXPORT
    st.subheader("📤 Export Laporan")
    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
        for sheet,d in df_lr.items(): d.to_excel(writer, sheet_name=sheet, index=False)
        df_aset.to_excel(writer, sheet_name="Aset", index=False)
        df_kewajiban.to_excel(writer, sheet_name="Kewajiban", index=False)
        df_ekuitas.to_excel(writer, sheet_name="Ekuitas", index=False)
    st.download_button("⬇️ Download Excel", data=excel_buf.getvalue(), file_name="laporan_keuangan.xlsx")

    pdf_lr = export_pdf_laba_rugi(df_lr, laba_rugi, nama_pt, pejabat, periode_text)
    st.download_button("⬇️ Download PDF Laba Rugi", data=pdf_lr, file_name="Laporan_Laba_Rugi.pdf", mime="application/pdf")

    pdf_nr = export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, pejabat, periode_text)
    st.download_button("⬇️ Download PDF Neraca", data=pdf_nr, file_name="Laporan_Posisi_Keuangan.pdf", mime="application/pdf")
