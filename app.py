import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

# ========== Fungsi utilitas ==========
def normalisasi_kolom(df: pd.DataFrame):
    df.columns = df.columns.str.strip().str.lower()
    rename_map = {
        "kode akun": "kode_akun",
        "kodeakun": "kode_akun",
        "nama akun": "nama_akun",
        "saldo awal": "saldo_awal",
        "saldo": "saldo_awal",
        "posisi normal akun": "posisi_normal_akun",
        "laporan": "laporan",
        "sub tipe laporan": "sub_tipe_laporan",
    }
    df.rename(columns={k:v for k,v in rename_map.items() if k in df.columns}, inplace=True)
    return df

def hitung_saldo(saldo_awal, debit, kredit, posisi):
    if str(posisi).lower() == "debit":
        return saldo_awal + debit - kredit
    else:
        return saldo_awal - debit + kredit

# ========== Fungsi export Laba Rugi ==========
def export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 3*cm

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 0.7*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, y, "LAPORAN LABA RUGI")
    y -= 0.6*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, f"Untuk Periode yang Berakhir Pada {periode_text}")
    y -= 1*cm

    def tulis_baris(text, amount=None, bold=False, total=False):
        nonlocal y
        if bold: c.setFont("Helvetica-Bold", 9)
        else: c.setFont("Helvetica", 9)
        c.drawString(2*cm, y, text)
        if amount is not None:
            c.drawRightString(w-2*cm, y, f"Rp {amount:,.0f}")
            if total:
                c.line(w-6*cm, y-2, w-2*cm, y-2)
        y -= 0.45*cm

    for _, r in df_laba.iterrows():
        if "header" in str(r["tipe_akun"]).lower():
            tulis_baris(r["nama_akun"], None, bold=True)
        else:
            tulis_baris("   " + r["nama_akun"], r["saldo_akhir_adj"])

    y -= 0.3*cm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "LABA (RUGI) BERSIH")
    c.drawRightString(w-2*cm, y, f"Rp {laba_bersih:,.0f}")
    c.line(w-6*cm, y-2, w-2*cm, y-2)
    c.line(w-6*cm, y-6, w-2*cm, y-6)

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# ========== Fungsi export Neraca ==========
def export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    y = h - 3*cm

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(w/2, y, nama_pt)
    y -= 0.7*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(w/2, y, "LAPORAN POSISI KEUANGAN")
    y -= 0.6*cm
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, y, f"Per {periode_text}")
    y -= 1*cm

    def tulis_baris(text, amount=None, bold=False, total=False):
        nonlocal y
        if bold: c.setFont("Helvetica-Bold", 9)
        else: c.setFont("Helvetica", 9)
        c.drawString(2*cm, y, text)
        if amount is not None:
            c.drawRightString(w-2*cm, y, f"Rp {amount:,.0f}")
            if total:
                c.line(w-6*cm, y-2, w-2*cm, y-2)
        y -= 0.45*cm

    def tulis_section(title, df_section, total):
        nonlocal y
        tulis_baris(title, None, bold=True)
        for _, r in df_section.iterrows():
            tulis_baris("   " + r["nama_akun"], r["saldo_akhir_adj"])
        tulis_baris(f"TOTAL {title.upper()}", total, bold=True, total=True)
        y -= 0.3*cm

    tulis_section("ASET", df_aset, total_aset)
    tulis_section("KEWAJIBAN", df_kewajiban, total_kewajiban)
    tulis_section("EKUITAS", df_ekuitas, total_ekuitas)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, "TOTAL KEWAJIBAN + EKUITAS")
    c.drawRightString(w-2*cm, y, f"Rp {total_kewajiban+total_ekuitas:,.0f}")
    c.line(w-6*cm, y-2, w-2*cm, y-2)
    c.line(w-6*cm, y-6, w-2*cm, y-6)

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# ========== Streamlit App ==========
st.title("📊 Aplikasi Laporan Keuangan")

coa_file = st.file_uploader("Upload COA.xlsx", type=["xlsx"])
saldo_file = st.file_uploader("Upload Saldo Awal.xlsx", type=["xlsx"])
jurnal_file = st.file_uploader("Upload Jurnal.xlsx", type=["xlsx"])
nama_pt = st.text_input("Nama Perusahaan", "PT Contoh Sejahtera")
tanggal_awal = st.date_input("Tanggal Awal")
tanggal_akhir = st.date_input("Tanggal Akhir")

if coa_file and saldo_file and jurnal_file:
    coa = normalisasi_kolom(pd.read_excel(coa_file))
    saldo_awal = normalisasi_kolom(pd.read_excel(saldo_file))
    jurnal = normalisasi_kolom(pd.read_excel(jurnal_file))

    # Pastikan kolom minimal ada
    if "kode_akun" not in coa.columns: st.stop()
    if "kode_akun" not in saldo_awal.columns: saldo_awal["kode_akun"] = coa["kode_akun"]
    if "saldo_awal" not in saldo_awal.columns: saldo_awal["saldo_awal"] = 0
    if "debit" not in jurnal.columns: jurnal["debit"] = 0
    if "kredit" not in jurnal.columns: jurnal["kredit"] = 0

    # Gabung
    df = coa.merge(saldo_awal[["kode_akun","saldo_awal"]], on="kode_akun", how="left").fillna(0)
    jurnal_sum = jurnal.groupby("kode_akun")[["debit","kredit"]].sum().reset_index()
    df = df.merge(jurnal_sum, on="kode_akun", how="left").fillna(0)

    df["saldo_akhir"] = df.apply(lambda r: hitung_saldo(r["saldo_awal"], r["debit"], r["kredit"], r["posisi_normal_akun"]), axis=1)

    # Atur saldo sesuai laporan
    def adjust_saldo(row):
        laporan = str(row["laporan"]).lower()
        sub = str(row["sub_tipe_laporan"]).lower()
        normal = str(row["posisi_normal_akun"]).lower()
        val = row["saldo_akhir"]
        if "posisi keuangan" in laporan:
            if "aset" in sub and normal == "debit": return val
            if "aset" in sub and normal == "kredit": return -val
            if "kewajiban" in sub and normal == "kredit": return val
            if "ekuitas" in sub and normal == "kredit": return val
            return -val
        elif "laba rugi" in laporan:
            return abs(val)
        return val

    df["saldo_akhir_adj"] = df.apply(adjust_saldo, axis=1)

    # Pisahkan laporan
    df_laba = df[df["laporan"].str.contains("laba", case=False, na=False)]
    df_aset = df[df["sub_tipe_laporan"].str.contains("aset", case=False, na=False)]
    df_kewajiban = df[df["sub_tipe_laporan"].str.contains("kewajiban", case=False, na=False)]
    df_ekuitas = df[df["sub_tipe_laporan"].str.contains("ekuitas", case=False, na=False)]

    # Laba bersih
    pendapatan = df_laba[df_laba["sub_tipe_laporan"].str.contains("pendapatan", case=False, na=False)]["saldo_akhir_adj"].sum()
    beban = df_laba[df_laba["sub_tipe_laporan"].str.contains("beban", case=False, na=False)]["saldo_akhir_adj"].sum()
    laba_bersih = pendapatan - beban

    # Tambah laba berjalan ke ekuitas
    df_ekuitas = pd.concat([
        df_ekuitas,
        pd.DataFrame([{"kode_akun":"3004","nama_akun":"Laba (Rugi) Berjalan","saldo_akhir_adj":laba_bersih}])
    ])

    total_aset = df_aset["saldo_akhir_adj"].sum()
    total_kewajiban = df_kewajiban["saldo_akhir_adj"].sum()
    total_ekuitas = df_ekuitas["saldo_akhir_adj"].sum()

    periode_text = tanggal_akhir.strftime("%d %B %Y")

    pdf_lr = export_pdf_laba_rugi(df_laba, laba_bersih, nama_pt, periode_text)
    st.download_button("⬇️ Download Laporan Laba Rugi (PDF)", pdf_lr, file_name="Laporan_Laba_Rugi.pdf")

    pdf_nr = export_pdf_neraca(df_aset, df_kewajiban, df_ekuitas, total_aset, total_kewajiban, total_ekuitas, nama_pt, periode_text)
    st.download_button("⬇️ Download Laporan Posisi Keuangan (PDF)", pdf_nr, file_name="Laporan_Posisi_Keuangan.pdf")

    # Excel export
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df_laba.to_excel(writer, sheet_name="Laba Rugi", index=False)
        df_aset.to_excel(writer, sheet_name="Aset", index=False)
        df_kewajiban.to_excel(writer, sheet_name="Kewajiban", index=False)
        df_ekuitas.to_excel(writer, sheet_name="Ekuitas", index=False)
        df.to_excel(writer, sheet_name="Gabungan", index=False)
    st.download_button("⬇️ Download Semua (Excel)", out.getvalue(), file_name="Laporan_Keuangan.xlsx")
