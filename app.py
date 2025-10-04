import streamlit as st
import pandas as pd
import io
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

st.set_page_config(page_title="Laporan Keuangan Profesional", layout="wide")
st.title("📊 Aplikasi Laporan Keuangan Profesional")

# ============= Upload files =============
uploaded_coa = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
uploaded_saldo = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
uploaded_jurnal = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])
tanggal_awal = st.date_input("Tanggal Awal Periode")
tanggal_akhir = st.date_input("Tanggal Akhir Periode")
nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")

# ============= Helper functions =============
def bersihkan_kolom(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df

def normalisasi_kode(df):
    kemungkinan = ["kodeakun", "kode_akun", "akun_kode", "akun", "no_akun", "rekening"]
    for k in kemungkinan:
        if k in df.columns:
            df.rename(columns={k: "kode_akun"}, inplace=True)
            return df
    df.rename(columns={df.columns[0]: "kode_akun"}, inplace=True)
    return df

def hitung_saldo(saldo, debit, kredit, normal):
    if str(normal).lower() == "debit":
        return saldo + debit - kredit
    else:
        return saldo - debit + kredit

# ============= PDF LABA RUGI =============
def buat_pdf_laba_rugi(df, laba_bersih, nama_pt, periode_text):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    y = h - 5 * cm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w / 2, h - 2 * cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w / 2, h - 3 * cm, "LAPORAN LABA RUGI")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w / 2, h - 3.7 * cm, f"Untuk Periode yang Berakhir Pada {periode_text}")
    c.rect(2 * cm, 2 * cm, w - 4 * cm, h - 6 * cm)

    def tulis(teks, nilai=None, tebal=False, garis=False):
        nonlocal y
        if y < 3 * cm:
            c.showPage()
            y = h - 2 * cm
        c.setFont("Helvetica-Bold" if tebal else "Helvetica", 10)
        c.drawString(2.5 * cm if nilai is None else 3 * cm, y, teks)
        if nilai not in [None, 0]:
            if garis:
                c.line(w - 5 * cm, y + 0.2 * cm, w - 2 * cm, y + 0.2 * cm)
            c.drawRightString(w - 2 * cm, y, f"Rp {nilai:,.0f}")
        y -= 0.5 * cm

    for sub, isi in df.groupby("sub_tipe_laporan"):
        tulis(sub, tebal=True)
        subtotal = 0
        for _, r in isi.iterrows():
            if "header" in str(r.get("tipe_akun", "")).lower():
                tulis(r.get("nama_akun", ""), tebal=True)
            elif r.get("saldo_akhir", 0) != 0:
                tulis("   " + r.get("nama_akun", ""), r.get("saldo_akhir", 0))
                subtotal += r.get("saldo_akhir", 0)
        tulis(f"TOTAL {sub}", subtotal, tebal=True, garis=True)
        y -= 0.3 * cm

    c.line(w - 5 * cm, y, w - 2 * cm, y)
    c.line(w - 5 * cm, y - 0.2 * cm, w - 2 * cm, y - 0.2 * cm)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2.5 * cm, y - 0.7 * cm, "LABA (RUGI) BERSIH")
    c.drawRightString(w - 2 * cm, y - 0.7 * cm, f"Rp {laba_bersih:,.0f}")
    c.save()
    buf.seek(0)
    return buf

# ============= PDF NERACA =============
def buat_pdf_neraca(df_aset, df_kewajiban, df_ekuitas,
                    total_aset, total_kewajiban, total_ekuitas,
                    nama_pt, periode_text):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    y = h - 5 * cm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w / 2, h - 2 * cm, nama_pt)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w / 2, h - 3 * cm, "LAPORAN POSISI KEUANGAN")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w / 2, h - 3.7 * cm, f"Per {periode_text}")
    c.rect(2 * cm, 2 * cm, w - 4 * cm, h - 6 * cm)

    def draw_sec(judul, df, total):
        nonlocal y
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2.5 * cm, y, judul.upper())
        y -= 0.4 * cm
        for _, r in df.iterrows():
            if r.get("saldo_akhir_adj", 0) != 0:
                c.setFont("Helvetica", 9)
                c.drawString(3 * cm, y, r.get("nama_akun", ""))
                c.drawRightString(w - 2 * cm, y, f"Rp {r['saldo_akhir_adj']:,.0f}")
                y -= 0.4 * cm
        c.line(w - 5 * cm, y + 0.1 * cm, w - 2 * cm, y + 0.1 * cm)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(3 * cm, y - 0.5 * cm, f"TOTAL {judul}")
        c.drawRightString(w - 2 * cm, y - 0.5 * cm, f"Rp {total:,.0f}")
        y -= 1 * cm

    draw_sec("ASET", df_aset, total_aset)
    draw_sec("KEWAJIBAN", df_kewajiban, total_kewajiban)
    draw_sec("EKUITAS", df_ekuitas, total_ekuitas)
    c.line(w - 5 * cm, y, w - 2 * cm, y)
    c.line(w - 5 * cm, y - 0.2 * cm, w - 2 * cm, y - 0.2 * cm)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2.5 * cm, y - 0.7 * cm, "TOTAL KEWAJIBAN + EKUITAS")
    c.drawRightString(w - 2 * cm, y - 0.7 * cm, f"Rp {(total_kewajiban + total_ekuitas):,.0f}")
    c.save()
    buf.seek(0)
    return buf

# ============= MAIN PROSES =============
if uploaded_coa and uploaded_saldo and uploaded_jurnal:
    coa = bersihkan_kolom(normalisasi_kode(pd.read_excel(uploaded_coa)))
    saldo_awal = bersihkan_kolom(normalisasi_kode(pd.read_excel(uploaded_saldo)))
    jurnal = bersihkan_kolom(pd.read_excel(uploaded_jurnal))

    # -------- deteksi kolom kode akun otomatis --------
    possible_names = ["kode_akun", "kodeakun", "akun_kode", "akun", "no_akun", "rekening"]
    col_kode = next((c for c in jurnal.columns if c in possible_names), None)
    if not col_kode:
        st.error("❌ Kolom kode akun tidak ditemukan di file Jurnal. Pastikan ada kolom seperti 'Kode Akun'.")
        st.stop()
    jurnal.rename(columns={col_kode: "kode_akun"}, inplace=True)

    for kol in ["debit", "kredit"]:
        if kol not in jurnal.columns:
            jurnal[kol] = 0

    jurnal["debit"] = pd.to_numeric(jurnal["debit"], errors="coerce").fillna(0)
    jurnal["kredit"] = pd.to_numeric(jurnal["kredit"], errors="coerce").fillna(0)
    jurnal["kode_akun"] = jurnal["kode_akun"].astype(str).str.strip()
    jurnal["kode_akun"] = jurnal["kode_akun"].str.replace(";", ",")
    if jurnal["kode_akun"].str.contains(",").any():
        jurnal = jurnal.assign(kode_akun=jurnal["kode_akun"].str.split(",")).explode("kode_akun")
        jurnal["kode_akun"] = jurnal["kode_akun"].str.strip()

    # -------- proses data utama --------
    df = coa.merge(saldo_awal, on="kode_akun", how="left").fillna(0)
    jurnal["tanggal"] = pd.to_datetime(jurnal.get("tanggal"), errors="coerce")
    mutasi = jurnal[(jurnal["tanggal"] >= pd.to_datetime(tanggal_awal)) &
                    (jurnal["tanggal"] <= pd.to_datetime(tanggal_akhir))]
    mutasi_group = mutasi.groupby("kode_akun", dropna=False)[["debit", "kredit"]].sum().reset_index()
    df = df.merge(mutasi_group, on="kode_akun", how="left").fillna(0)

    for kolom in ["nama_akun", "tipe_akun", "posisi_normal_akun", "laporan", "sub_tipe_laporan"]:
        if kolom not in df.columns:
            df[kolom] = ""

    df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(
        r.get("saldo_awal", 0), r.get("debit", 0), r.get("kredit", 0),
        r.get("posisi_normal_akun", "debit")), axis=1)
    df["saldo_akhir_adj"] = df.apply(lambda r:
        r["saldo_akhir"] if r["posisi_normal_akun"].lower() == "debit"
        else -r["saldo_akhir"] if r["saldo_akhir"] < 0 else r["saldo_akhir"], axis=1)

    # -------- pisah laporan --------
    df_laba = df[df["laporan"].str.contains("laba", case=False, na=False)]
    df_neraca = df[df["laporan"].str.contains("posisi", case=False, na=False)]

    laba_bersih = df_laba["saldo_akhir_adj"].sum()
    df_aset = df_neraca[df_neraca["sub_tipe_laporan"].str.contains("aset", case=False, na=False)]
    df_kewajiban = df_neraca[df_neraca["sub_tipe_laporan"].str.contains("kewajiban", case=False, na=False)]
    df_ekuitas = df_neraca[df_neraca["sub_tipe_laporan"].str.contains("ekuitas", case=False, na=False)].copy()
    df_ekuitas.loc[df_ekuitas["nama_akun"].str.contains("laba", case=False, na=False), "saldo_akhir_adj"] = laba_bersih

    total_aset = df_aset["saldo_akhir_adj"].sum()
    total_kewajiban = df_kewajiban["saldo_akhir_adj"].sum()
    total_ekuitas = df_ekuitas["saldo_akhir_adj"].sum()

    periode_text = tanggal_akhir.strftime("%d %B %Y")
    pdf_lr = buat_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text)
    pdf_nr = buat_pdf_neraca(df_aset, df_kewajiban, df_ekuitas,
                             total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text)

    st.success("✅ Laporan berhasil dibuat.")
    st.download_button("⬇️ Laporan Laba Rugi (PDF)", data=pdf_lr, file_name="Laba_Rugi.pdf", mime="application/pdf")
    st.download_button("⬇️ Laporan Posisi Keuangan (PDF)", data=pdf_nr, file_name="Posisi_Keuangan.pdf", mime="application/pdf")

    # Export Excel
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
        df_laba.to_excel(writer, sheet_name="Laba_Rugi", index=False)
        df_neraca.to_excel(writer, sheet_name="Neraca", index=False)
    st.download_button("⬇️ Export ke Excel", data=excel_buf.getvalue(),
                       file_name="Laporan_Keuangan.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
