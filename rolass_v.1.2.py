from flask import Flask, render_template_string, request
import requests
import pandas as pd
import plotly.express as px
import plotly
import json
from flask import send_file
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import re
from plotly.utils import PlotlyJSONEncoder
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import Alignment, Border, Side
from openpyxl.styles import PatternFill
import os, tempfile
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib import colors
from reportlab.pdfbase.ttfonts import TTFont
import locale

app = Flask(__name__)

# Bersihkan string agar aman untuk nama file
def safe_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "_", str(text).strip())

def load_data():
    url = "https://rol.postel.go.id/api/observasi/allapproved"
    params = {
        "upt": 19,
        "year": "2025",
        "pageIndex": 1,
        "pageSize": 100000
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "*/*",
        "Referer": "https://rol.postel.go.id/observasi/laporan",
        "X-Requested-With": "XMLHttpRequest",
        "Cookie": "csrf_cookie_name=6701fbfe229f77d47c710eec3391f386; ci_session=1tmigk58v1636foa82v97rmo8o2olsrc"
    }

    try:
        response = requests.get(url, params=params, headers=headers)
        data = response.json().get("data", [])
        return pd.DataFrame(data)
    except Exception as e:
        print("Gagal ambil data API:", e)
        return pd.DataFrame()

def load_info_inspeksi():
    url = "https://apstard.postel.go.id/dashboard/info-inspeksi-3"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://apstard.postel.go.id/dashboard/dashboard-keseluruhan-upt",
        # 🔴 Cookie perlu diganti sesuai hasil loginmu
        "Cookie": "csrf_cookie_name=130a5f593c00deef94f76eff425ae17c; ci_session=a70s64f63e2tkmdfm0s2dab9mb1f22ho",
    }

    payload = {
        "periode": "2025",
        "upt_id": "14"
    }

    r = requests.post(url, headers=headers, data=payload)
    if r.status_code != 200:
        print("Request gagal:", r.status_code, r.text)
        return pd.DataFrame()

    data = r.json()
    print("✅ Response diterima:", data)

    # Convert ke DataFrame
    df = pd.DataFrame([data])
    return df


def load_pantib():
    url = "https://rol.postel.go.id/api/penertiban/list"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://rol.postel.go.id/penertiban",
        # ⚠️ Cookie harus diganti manual setiap login
        "Cookie": "csrf_cookie_name=4d68a695308880e740aed8eb9ddaf59b; ci_session=562m00244moanb6eb57e0malktpao9vb"
    }

    all_data = []
    page = 1
    page_size = 10000  

    while True:
        params = {
            "status": "",
            "status_penertiban": "",
            "tahun": "2025",
            "pageIndex": page,
            "pageSize": page_size
        }
        r = requests.get(url, headers=headers, params=params)
        if r.status_code != 200:
            print(f"Request gagal di page {page}: {r.status_code}")
            break

        data = r.json().get("data", [])
        if not data:
            break
        all_data.extend(data)
        page += 1

    return pd.DataFrame(all_data)


def format_tanggal_indonesia(tgl_raw):
    bulan_map = {
        "January": "Januari", "February": "Februari", "March": "Maret",
        "April": "April", "May": "Mei", "June": "Juni",
        "July": "Juli", "August": "Agustus", "September": "September",
        "October": "Oktober", "November": "November", "December": "Desember"
    }

    try:
        # parse string YYYY-MM-DD
        dt = datetime.strptime(tgl_raw, "%Y-%m-%d")
        # cek OS (Windows pakai %#d, Linux pakai %-d)
        try:
            day_str = dt.strftime("%-d")   # Linux/Unix
        except:
            day_str = dt.strftime("%#d")  # Windows
        month_str = dt.strftime("%B")
        year_str = dt.strftime("%Y")

        # ganti bulan ke bahasa Indonesia
        month_str = bulan_map.get(month_str, month_str)

        return f"{day_str} {month_str} {year_str}"
    except Exception:
        return tgl_raw  # fallback kalau parsing gagal


@app.route("/unduh_laporan", methods=["GET", "POST"])
def unduh_laporan():       
    df = load_data()

    # Transformasi jenis identifikasi
    df["observasi_status_identifikasi_name"] = df["observasi_status_identifikasi_name"].str.replace(
        r"OFF AIR \(Sedang Tidak Digunakan\)", "OFF AIR", regex=True
    )
    df['jenis'] = df['observasi_status_identifikasi_name'].apply(
        lambda x: 'Belum Teridentifikasi' if x == 'BELUM DIKETAHUI' else 'Teridentifikasi'
    )
    
    locale.setlocale(locale.LC_TIME, "id_ID.utf8")  # aktifkan format tanggal bahasa Indonesia
    tanggal_skrg = datetime.now().strftime("%d %B %Y")

    # Ambil filter dari form (POST) atau query (GET)
    if request.method == "POST":
        selected_spt = request.form.get("spt", "Semua")
        selected_kab = request.form.get("kab", "Semua")
        selected_kec = request.form.get("kec", "Semua")
        selected_cat = request.form.get("cat", "Semua")
        pelaksana_list = request.form.getlist("pelaksana")
        perangkat = request.form.get("perangkat", "Tetap/Transportable TCI")
        tgl_spt_raw = request.form.get("tgl_spt")
        if tgl_spt_raw:
                tgl_spt = format_tanggal_indonesia(tgl_spt_raw)
        else:
                tgl_spt = format_tanggal_indonesia(datetime.now().strftime("%Y-%m-%d"))
        
    else:
        selected_spt = request.args.get("spt", "Semua")
        selected_kab = request.args.get("kab", "Semua")
        selected_kec = request.args.get("kec", "Semua")
        selected_cat = request.args.get("cat", "Semua")
        pelaksana_list = []
        tgl_spt = format_tanggal_indonesia(datetime.now().strftime("%Y-%m-%d"))
        perangkat = "Tetap/Transportable TCI"

    # Filter data
    filt = df.copy()
    if selected_spt != "Semua":
        filt = filt[filt["observasi_no_spt"] == selected_spt]
    if selected_kab != "Semua":
        filt = filt[filt["observasi_kota_nama"] == selected_kab]
    if selected_kec != "Semua":
        filt = filt[filt["observasi_kecamatan_nama"] == selected_kec]
    if selected_cat != "Semua":
        filt = filt[filt["scan_catatan"] == selected_cat]

    # pastikan kolom tanggal dalam bentuk datetime
    filt["observasi_tanggal"] = pd.to_datetime(filt["observasi_tanggal"], errors="coerce")
    
    # ambil bulan & tahun dari data observasi
    if not filt.empty and filt["observasi_tanggal"].notna().any():
        bulan_tahun_obs = filt["observasi_tanggal"].dt.strftime("%B %Y").iloc[0]
    else:
        bulan_tahun_obs = datetime.now().strftime("%B %Y")


    # === Registrasi font ===
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
    #pdfmetrics.registerFont(TTFont("BrushScript", "BRUSHSCI.ttf"))
    pdfmetrics.registerFont(TTFont("BrushScript", r"C:\Users\BALMON_MATARAM\ROL-Assistance-Summary-System\BRUSHSCI.ttf"))
    pdfmetrics.registerFont(TTFont("zph", r"C:\Users\BALMON_MATARAM\ROL-Assistance-Summary-System\bodoni-six-itc-bold-italic-os-5871d33e4dc4a.ttf"))
    pdfmetrics.registerFont(TTFont("Arial", r"C:\Users\BALMON_MATARAM\ROL-Assistance-Summary-System\ARIALBD.ttf"))
    pdfmetrics.registerFont(TTFont("Arialbd", r"C:\Users\BALMON_MATARAM\ROL-Assistance-Summary-System\ARIAL.ttf"))

    # Simpan sementara
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    filename = tmp.name
    doc = SimpleDocTemplate(filename, pagesize=A4,
                            rightMargin=50, leftMargin=50,
                            topMargin=20, bottomMargin=30)

    styles = getSampleStyleSheet()
    style_center = ParagraphStyle(name="Center", parent=styles["Normal"], alignment=TA_CENTER, fontName="Arial", fontSize=14, leading=15)
    style_center2 = ParagraphStyle(name="Center", parent=styles["Normal"], alignment=TA_CENTER, fontName="Arialbd", fontSize=12, leading=15)
    style_normal = ParagraphStyle(name="Normal", parent=styles["Normal"], alignment=TA_JUSTIFY, fontName="Arialbd", fontSize=12, leading=14)

    style_left_h1b = ParagraphStyle(name="Left", parent=styles["Normal"], alignment=TA_LEFT, fontName="zph", fontSize=16, leading=16, textColor=colors.blue)
    style_left_h2b = ParagraphStyle(name="Left", parent=styles["Normal"], alignment=TA_LEFT, fontName="BrushScript", fontSize=14, leading=16, textColor=colors.blue)
    style_left_h3b = ParagraphStyle(name="Left", parent=styles["Normal"], alignment=TA_LEFT, fontName="zph", fontSize=11, leading=16, textColor=colors.blue)
    style_left_h4b = ParagraphStyle(name="Left", parent=styles["Normal"], alignment=TA_LEFT, fontName="zph", fontSize=9, leading=16, textColor=colors.blue)
    style_left_h4t = ParagraphStyle(name="Left", parent=styles["Normal"], alignment=TA_LEFT, fontName="Arialbd", fontSize=12, leading=14)

    content = []
    

    # === Kop Surat ===
    logo_path = os.path.join(app.static_folder, "logo-kominfo.png")
    logo = Image(logo_path, width=70, height=70)
    kop_text = [
        Paragraph("<b>KEMENTERIAN KOMUNIKASI DAN INFORMATIKA RI</b>", style_left_h1b),
        Paragraph("DIREKTORAT JENDERAL SUMBER DAYA DAN PERANGKAT POS DAN INFORMATIKA", style_left_h3b),
        Paragraph("BALAI MONITOR SPEKTRUM FREKUENSI RADIO KELAS II MATARAM", style_left_h3b),
        Paragraph("Indonesia Terkoneksi : Makin Digital, Makin Maju", style_left_h2b),
        Paragraph("Jl.Singosari No.4 Mataram 83127 Telp.(0370) 646411 Fax.(0370) 648740-42 email: upt_mataram.postel.go.id", style_left_h4b)
    ]
    kop_table = Table([[logo, kop_text]], colWidths=[70, 430])
    kop_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LINEBELOW", (0,0), (-1,-1), 2, colors.blue)
    ]))
    content.append(kop_table)
    content.append(Spacer(1, 20))        
        
    # === Judul ===
    content.append(Paragraph("<b>NOTA DINAS</b>", style_center))
    # ambil bulan/tahun sekarang
    bulan_tahun = datetime.now().strftime("%m/%Y")
    content.append(Paragraph(f"<b>Nomor :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;/ND/Montib/{bulan_tahun}</b>", style_center2))
    content.append(Spacer(1, 20))

    # === Tabel Yth ===
    yth_text = [
        Paragraph("Yth", style_left_h4t),
        Paragraph("Dari", style_left_h4t),
        Paragraph("Hal", style_left_h4t),
        Paragraph("Sifat", style_left_h4t),
        Paragraph("Lampiran", style_left_h4t),
        Paragraph("Tanggal", style_left_h4t)
    ]
    yth2_text = [
        Paragraph("Kepala Balai Monitor SFR Kelas II Mataram", style_normal),
        Paragraph("Ketua Tim Kerja Monitoring dan Penertiban SFR dan APT", style_normal),
        Paragraph(f"Laporan Pelaksanaan Kegiatan Monitoring dan Identifikasi 15 Pita Frekuensi Radio dengan Perangkat SMFR {perangkat} Site {selected_kec}, {selected_kab} Bulan Agustus Tahun 2025", style_normal),
        Paragraph("Biasa", style_left_h4t),
        Paragraph("Satu bendel", style_left_h4t),
        Paragraph(f"{tanggal_skrg}", style_left_h4t)
    ]
    data = [[l, ":", r] for l, r in zip(yth_text, yth2_text)]
    yth_table = Table(data, colWidths=[70, 10, 420])
    yth_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),   # rata atas
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),   # tetap rata kiri
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    content.append(yth_table)
    content.append(Spacer(1, 20))
    
    # === Isi ===
    isi = f"""
    Dengan hormat disampaikan, bahwa berdasarkan Surat Tugas Nomor : 
    {selected_spt}, tanggal {tgl_spt}, tentang Kegiatan 
    Monitoring/Observasi dan Identifikasi 15 Pita Frekuensi Radio dengan Pemanfaatan 
    Perangkat SMFR {perangkat} Bulan {bulan_tahun_obs}, 
    terlampir kami sampaikan laporan pelaksanaan kegiatan dimaksud untuk 
    SMFR {perangkat} Site {selected_kec}, {selected_kab}.<br/><br/>
    
    Demikian disampaikan, mohon arahan lebih lanjut dan atas perhatian Bapak 
    diucapkan terimakasih.
    """
    content.append(Paragraph(isi, style_normal))
    content.append(Spacer(1, 80))
    

        
    # === TTD ===
    style_center_block = ParagraphStyle(name="CenterBlock", parent=styles["Normal"], alignment=TA_CENTER, fontName="Arialbd", fontSize=12)
    
    pelaksana_list = request.form.getlist("pelaksana")
    
    pelaksana_table = Table(
        [["", Paragraph("Abdy Budiman Djara", style_center_block)]],
        colWidths=[350,150]  # sesuaikan lebar halaman
    )
    pelaksana_table.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "RIGHT")
    ]))
    content.append(pelaksana_table)
    content.append(Spacer(1, 50))        
    
    # Tambah pemisah halaman
    content.append(PageBreak())
    
    # === Kop Surat ===
    logo = Image("logo-kominfo.png", width=70, height=70)
    kop_text = [
        Paragraph("<b>KEMENTERIAN KOMUNIKASI DAN INFORMATIKA RI</b>", style_left_h1b),
        Paragraph("DIREKTORAT JENDERAL SUMBER DAYA DAN PERANGKAT POS DAN INFORMATIKA", style_left_h3b),
        Paragraph("BALAI MONITOR SPEKTRUM FREKUENSI RADIO KELAS II MATARAM", style_left_h3b),
        Paragraph("Indonesia Terkoneksi : Makin Digital, Makin Maju", style_left_h2b),
        Paragraph("Jl.Singosari No.4 Mataram 83127 Telp.(0370) 646411 Fax.(0370) 648740-42 email: upt_mataram.postel.go.id", style_left_h4b)
    ]
    kop_table = Table([[logo, kop_text]], colWidths=[70, 430])
    kop_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LINEBELOW", (0,0), (-1,-1), 2, colors.blue)
    ]))
    content.append(kop_table)
    content.append(Spacer(1, 20))        
        
    # === Judul ===
    content.append(Paragraph("<b>NOTA DINAS</b>", style_center))
    content.append(Spacer(1, 20))

    # === Tabel Yth ===
    yth_text = [
        Paragraph("Yth", style_left_h4t),
        Paragraph("Dari", style_left_h4t),
        Paragraph("Hal", style_left_h4t),
        Paragraph("Sifat", style_left_h4t),
        Paragraph("Tanggal", style_left_h4t)
    ]
    yth2_text = [
        Paragraph("Ketua Tim Kerja Monitoring dan Penertiban SFR dan APT", style_normal),
        Paragraph(f"Pelaksana Kegiatan Monitoring dan Identifikasi 15 Pita Frekuensi Radio dengan Perangkat SMFR {perangkat} Site {selected_kec}, {selected_kab}", style_normal),
        Paragraph(f"Laporan Pelaksanaan Kegiatan Monitoring dan Identifikasi 15 Pita Frekuensi Radio dengan Perangkat SMFR {perangkat} Site {selected_kec}, {selected_kab} Bulan Agustus Tahun 2025", style_normal),
        Paragraph("Biasa", style_left_h4t),
        Paragraph(f"{tanggal_skrg}", style_left_h4t)
    ]
    data = [[l, ":", r] for l, r in zip(yth_text, yth2_text)]
    yth_table = Table(data, colWidths=[70, 10, 420])
    yth_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),   # rata atas
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),   # tetap rata kiri
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    content.append(yth_table)
    content.append(Spacer(1, 20))
    
    # === Isi ===
    isi = f"""
    Dengan hormat disampaikan, bahwa berdasarkan Surat Tugas Nomor : 
    {selected_spt}, tanggal {tgl_spt}, tentang Kegiatan 
    Monitoring/Observasi dan Identifikasi 15 Pita Frekuensi Radio dengan Pemanfaatan 
    Perangkat SMFR {perangkat} Bulan {bulan_tahun_obs}, 
    terlampir kami sampaikan laporan pelaksanaan kegiatan dimaksud untuk 
    SMFR {perangkat} Site {selected_kec}, {selected_kab}.<br/><br/>
    
    Demikian disampaikan, mohon arahan lebih lanjut dan atas perhatian Bapak 
    diucapkan terimakasih.
    """
    content.append(Paragraph(isi, style_normal))
    content.append(Spacer(1, 80))
        
    # === TTD ===
    style_center_block = ParagraphStyle(name="CenterBlock", parent=styles["Normal"], alignment=TA_CENTER, fontName="Arialbd", fontSize=12)
    
    pelaksana_list = request.form.getlist("pelaksana")
    
    for p in pelaksana_list:
        pelaksana_table = Table(
            [["", Paragraph(p, style_center_block)]],
            colWidths=[350,150]  # sesuaikan lebar halaman
        )
        pelaksana_table.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "RIGHT")
        ]))
        content.append(pelaksana_table)
        content.append(Spacer(1, 50))
    
    # Build PDF
    doc.build(content)
    
    return send_file(filename, as_attachment=True, download_name="Nota_Dinas.pdf")

@app.route("/download_excel", methods=["POST"])
def download_excel():
    df = load_data()

    # Transformasi jenis identifikasi
    df["observasi_status_identifikasi_name"] = df["observasi_status_identifikasi_name"].str.replace(
        r"OFF AIR \(Sedang Tidak Digunakan\)", "OFF AIR", regex=True
    )
    df['jenis'] = df['observasi_status_identifikasi_name'].apply(
        lambda x: 'Belum Teridentifikasi' if x == 'BELUM DIKETAHUI' else 'Teridentifikasi'
    )

    # Ambil filter dari form
    selected_spt = request.form.get("spt", "Semua")
    selected_kab = request.form.get("kab", "Semua")
    selected_kec = request.form.get("kec", "Semua")
    selected_cat = request.form.get("cat", "Semua")

    # Filter data
    filt = df.copy()
    if selected_spt != "Semua":
        filt = filt[filt["observasi_no_spt"] == selected_spt]
    if selected_kab != "Semua":
        filt = filt[filt["observasi_kota_nama"] == selected_kab]
    if selected_kec != "Semua":
        filt = filt[filt["observasi_kecamatan_nama"] == selected_kec]
    if selected_cat != "Semua":
        filt = filt[filt["scan_catatan"] == selected_cat]

    if filt.empty:
        return "<h3>Data kosong, tidak bisa disimpan.</h3>"

    # Ringkasan
    jumlah_kota = filt['observasi_kota_nama'].nunique()
    legal = filt.groupby(['observasi_kota_nama', 'observasi_status_identifikasi_name']).size().reset_index(name='Jumlah')
    band = filt.pivot_table(index=['observasi_kota_nama', 'band_nama', 'jenis'], aggfunc='size', fill_value=0)
    dinas = filt.pivot_table(index=['observasi_kota_nama', 'observasi_service_name', 'jenis'], aggfunc='size', fill_value=0)
    pita = filt.pivot_table(index=['observasi_kota_nama', 'observasi_range_frekuensi', 'observasi_status_identifikasi_name'], aggfunc='size', fill_value=0)

    try:
        # === Load ISR & Samakan Format ===
        base_dir = os.path.dirname(os.path.abspath(__file__))
        isr_path = os.path.join(base_dir, "Data Target Monitor ISR 2025 - Mataram.csv")
    
        # Load CSV
        df_ISR = pd.read_csv(isr_path, on_bad_lines='skip', delimiter=';') \
                    .rename(columns={'Freq': 'Frekuensi', 'Clnt Name': 'Identifikasi'})
        df_ISR['Frekuensi'] = pd.to_numeric(df_ISR['Frekuensi'], errors='coerce')
        filt['observasi_frekuensi'] = pd.to_numeric(filt['observasi_frekuensi'], errors='coerce')
    
        # Filter ISR bila kab dipilih
        if selected_kab != "Semua":
            df_ISR = df_ISR[df_ISR['Kab/Kota'].astype(str).str.strip().str.upper() == selected_kab.strip().upper()]
    
        # === Hitung kesesuaian dengan ISR ===
        freq_df1 = filt.groupby(
            ['observasi_frekuensi','observasi_sims_client_name','observasi_kota_nama']
        ).size().reset_index(name='Jumlah_df1')
    
        freq_df1 = freq_df1.rename(columns={
            'observasi_frekuensi': 'Frekuensi',
            'observasi_sims_client_name': 'Identifikasi',
            'observasi_kota_nama': 'Kab/Kota'
        })

        freq_df2 = df_ISR.groupby(['Frekuensi','Identifikasi','Kab/Kota']).size().reset_index(name='Jumlah_df2')
        
        # Hapus duplikat, hanya sisakan satu per kombinasi Frekuensi & Identifikasi
        merged = pd.merge(freq_df1, freq_df2, on=['Frekuensi', 'Identifikasi','Kab/Kota'], how='inner')
        
        jumlah_sesuai_isr = len(merged)

        # Ambil kota yang termonitor dari data observasi
        kota_termonitor = filt['observasi_kota_nama'].dropna().astype(str).str.strip().str.upper().unique()

        # Filter target ISR berdasarkan kota termonitor
        target_match = df_ISR[df_ISR['Kab/Kota'].str.strip().str.upper().isin(kota_termonitor)]
    
        # Hitung target ISR total (jumlah baris, karena tidak ada kolom "Jumlah ISR")
        jumlah_target_isr = len(target_match) if not target_match.empty else 0

        # Hitung persen kesesuaian
        persen_sesuai_isr = round((jumlah_sesuai_isr / jumlah_target_isr * 100), 2) if jumlah_target_isr > 0 else 0
    
    except Exception as e:
        print("Gagal hitung kesesuaian ISR:", e)
        jumlah_sesuai_isr, jumlah_target_isr, persen_sesuai_isr = 0, 0, 0



    jumlah_iden = len(filt[filt['jenis'] == 'Teridentifikasi'])
    jumlah_total = len(filt)
    persen_teridentifikasi = round((jumlah_iden / jumlah_total * 100), 2) if jumlah_total > 0 else 0

    # Buat Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Ringkasan Laporan"

    def add_centered_row(value, merge_cells=4):
        ws.merge_cells(start_row=ws.max_row + 1, start_column=1, end_row=ws.max_row + 1, end_column=merge_cells)
        cell = ws.cell(row=ws.max_row, column=1, value=value)
        cell.alignment = Alignment(horizontal='center')

    def add_labeled_row(label, value):
        ws.append([label, value])
    
    # Definisi border tipis
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Definisi border tipis
    thin_border2 = Border(
        bottom=Side(style='thin')
    )
            
    # Set orientasi dan scaling untuk siap print
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    
    # Scale agar muat halaman (70% dari ukuran normal)
    ws.page_setup.scale = 75
    
    # Atur margin
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

    # Mulai isi
    add_centered_row("RANGKUMAN HASIL LAPORAN SURAT TUGAS")
    add_centered_row(selected_spt)
    add_centered_row(f"{selected_kec} - {selected_kab}")
    add_centered_row("=" * 60)

    add_labeled_row("Jumlah Kab/Kota Termonitor:", jumlah_kota)
    add_centered_row("=" * 60)

    # Legalitas
    ws.append(["Rangkuman Berdasarkan Legalitas:"])
    for _, row in legal.iterrows():
        ws.append(row.tolist())
    add_centered_row("=" * 60)

    # Band
    ws.append(["Rangkuman Berdasarkan Band:"])
    band_df = band.reset_index()
    for _, row in band_df.iterrows():
        ws.append(row.tolist())
    add_centered_row("=" * 60)

    # Dinas
    ws.append(["Rangkuman Berdasarkan Dinas:"])
    dinas_df = dinas.reset_index()
    for _, row in dinas_df.iterrows():
        ws.append(row.tolist())
    add_centered_row("=" * 60)

    # Pita
    ws.append(["Rangkuman Berdasarkan Pita:"])
    pita_df = pita.reset_index()
    for _, row in pita_df.iterrows():
        ws.append(row.tolist())
    add_centered_row("=" * 60)

    # Tambahan ringkasan
    add_labeled_row("Jumlah Data Rekap:", jumlah_total)
    add_labeled_row("Jumlah Stasiun Radio Teridentifikasi:", f"{jumlah_iden} ({persen_teridentifikasi}%)")
    add_labeled_row("Jumlah Stasiun Radio Sesuai ISR:", f"{jumlah_sesuai_isr} ({persen_sesuai_isr}%)")

    add_centered_row("=" * 60)
    
    # Atur lebar kolom otomatis
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_len + 2
            
    # === Tambahkan Data Observasi Terfilter di Sheet Baru ===
    filt_sheet2 = filt.copy()

    # Urutkan berdasarkan observasi_id
    if "observasi_id" in filt_sheet2.columns:
        filt_sheet2 = filt_sheet2.sort_values(by="observasi_frekuensi", ascending=True)

    # Tambahkan kolom "No" mulai dari 1
    filt_sheet2.insert(0, "No", range(1, len(filt_sheet2) + 1))

    # Ubah kolom Pita Frekuensi → ambil sebelum "."
    #if "observasi_range_frekuensi" in filt_sheet2.columns:
        #filt_sheet2["observasi_range_frekuensi"] = filt_sheet2["observasi_range_frekuensi"].astype(str).str.split(".").str[0]

    # Mapping nama kolom
    rename_map = {
        "observasi_tanggal": "Tanggal",
        "observasi_jam": "Jam",
        "band_nama": "Band",
        "observasi_range_frekuensi": "Pita Frekuensi",
        "observasi_frekuensi": "Frekuensi",
        "observasi_level": "Level",
        "observasi_service_name": "Dinas",
        "observasi_subservice_name": "Subservice",
        "observasi_emisi_name": "Kelas Emisi",
        "observasi_equip_name": "Kelas Stasiun",
        "observasi_sims_client_name": "Nama Klien",
        "observasi_status_identifikasi_name": "Legalitas",
        "observasi_kelurahan_nama": "Kel/Desa",
        "observasi_kecamatan_nama": "Kecamatan",
        "observasi_kota_nama": "Kab/Kota",
        "observasi_propinsi_nama": "Provinsi",
        "observasi_scan_detail_lat": "Latitude",
        "observasi_scan_detail_long": "Longitude"
    }
    filt_sheet2 = filt_sheet2.rename(columns=rename_map)
    # Urutkan kolom sesuai urutan di rename_map
    ordered_cols = list(rename_map.values())
    filt_sheet2 = filt_sheet2[ordered_cols]

    # Hapus kolom tidak perlu
    drop_cols = [
        "observasi_jenis_perangkat","upt_nama","observasi_pita_frekuensi_name",
        "scan_catatan","observasi_no_spt","observasi_azimuth","observasi_jenis_stasiun",
        "observasi_tgl_spt","observasi_keterangan","observasi_status_request_delete",
        "observasi_radius","sims_tgl_query","sims_area_of_service","sims_station_name","jenis","observasi_id"
    ]
    filt_sheet2 = filt_sheet2.drop(columns=[c for c in drop_cols if c in filt_sheet2.columns], errors="ignore")

    # Buat sheet baru
    ws2 = wb.create_sheet(title="Data Terfilter")

    # Tulis header
    ws2.append(list(filt_sheet2.columns))
    
    # Tulis data
    for _, row in filt_sheet2.iterrows():
        ws2.append(row.tolist())

    # Atur lebar kolom otomatis
    for col in ws2.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws2.column_dimensions[col_letter].width = max_len + 2
    
    # Terapkan wrap text, rata tengah, dan border ke semua cell
    for row in ws2.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            
    # Persempit kolom tertentu menjadi setengah
    col_map = {v: k for k, v in enumerate(filt_sheet2.columns, start=1)}
    special_cols = ["Kelas Emisi", "Kelas Stasiun", "Nama Klien", "Pita Frekuensi"]
    
    for sc in special_cols:
        if sc in col_map:
            col_letter = get_column_letter(col_map[sc])
            current_width = ws2.column_dimensions[col_letter].width
            ws2.column_dimensions[col_letter].width = max(10, current_width / 2)
    
    # Set orientasi dan scaling untuk siap print
    ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
    
    # Scale agar muat halaman (70% dari ukuran normal)
    ws2.page_setup.scale = 50
    
    # Atur margin
    ws2.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)
    
    # Tambahkan print titles (header row selalu tampil di atas setiap halaman)
    ws2.print_title_rows = "1:1"

    # Fill abu-abu untuk header baris pertama
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    for cell in ws2[1]:  # row pertama = header
        cell.fill = header_fill

    # Mode page break preview
    ws.sheet_view.view = "pageBreakPreview"
    ws2.sheet_view.view = "pageBreakPreview"

    # Simpan ke file sementara
    from tempfile import NamedTemporaryFile
    tmp = NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.seek(0)

    # Ubah nama berdasarkan input
    safe_spt = safe_filename(selected_spt)
    safe_kab = safe_filename(selected_kab)
    safe_kec = safe_filename(selected_kec)
    filename = f"Rekap {safe_spt} {safe_kec} {safe_kab}.xlsx"

    return send_file(
        tmp.name,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/", methods=["GET", "POST"])
def index():
    # === Load Info Inspeksi ===
    df_inspeksi = load_info_inspeksi()
    
    if not df_inspeksi.empty:
        total_inspeksi = int(df_inspeksi["total_inspeksi"].iloc[0].replace(",", ""))
        sudah_inspeksi = int(df_inspeksi["sudah_inspeksi"].iloc[0].replace(",", ""))
        belum_inspeksi = int(df_inspeksi["belum_inspeksi"].iloc[0].replace(",", ""))
        sesuai = int(df_inspeksi["sesuai"].iloc[0].replace(",", ""))
        tidak_sesuai = int(df_inspeksi["tidak_sesuai"].iloc[0].replace(",", ""))
        ilegal = int(df_inspeksi["ilegal"].iloc[0].replace(",", ""))
        off_air = int(df_inspeksi["off_air"].iloc[0].replace(",", ""))
    
        capaian_inspeksi = round((sudah_inspeksi / total_inspeksi * 100), 2) if total_inspeksi > 0 else 0
    
        # Data untuk chart
        pie_inspeksi_status = pd.DataFrame({
            "Status": ["Sesuai", "Tidak Sesuai", "Ilegal", "Off Air"],
            "Jumlah": [sesuai, tidak_sesuai, ilegal, off_air]
        })
    
        # Data untuk chart
        bar_inspeksi_status = pd.DataFrame({
            "Status": ["Tidak Sesuai", "Ilegal"],
            "Jumlah": [tidak_sesuai, ilegal]
        })
    
        pie_inspeksi = px.pie(
            pie_inspeksi_status, names="Status", values="Jumlah",
            title="Distribusi Hasil Inspeksi",
            hole=0.4,
            color_discrete_sequence=["#006db0", "#00ade6", "#edbc1b", "#8f181b", "#EF4444", "#6B7280",
                                     "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                     "#044065", "#d5ad2b", "#884a4c"]
        )
        pie_inspeksi.update_layout(legend=dict(orientation="v", x=1, y=0.5))
    
        bar_inspeksi = px.bar(
            bar_inspeksi_status, x="Status", y="Jumlah", color="Status",
            title="Jumlah Pelanggaran Inspeksi",
            color_discrete_sequence=["#edbc1b", "#8f181b", "#006db0", "#00ade6", "#EF4444", "#6B7280",
                                     "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                     "#044065", "#d5ad2b", "#884a4c"]
        )
    else:
        total_inspeksi = sudah_inspeksi = capaian_inspeksi = 0
        pie_inspeksi = bar_inspeksi = {}
    
    
    ############PENERTIBANNNNNNNNNN
    df_pantib = load_pantib()
    # Card 1: jumlah pelanggaran
    jumlah_pelanggaran = len(df_pantib["no"].dropna())
    
    # Card 2: persentase telah ditertibkan
    total_data = len(df_pantib)
    sudah_ditertibkan = df_pantib["penertiban_no_penindakan"].notna().sum()
    persentase_ditertibkan = round((sudah_ditertibkan / total_data) * 100, 2) if total_data > 0 else 0

    df = load_data()
    # Tambahkan kolom jenis
    df['jenis'] = df['observasi_status_identifikasi_name'].apply(
        lambda x: 'Belum Teridentifikasi' if x == 'BELUM DIKETAHUI' else 'Teridentifikasi'
    )

    if df.empty:
        return "Data tidak tersedia. Periksa koneksi API atau cookie."

    # Ambil pilihan dari form
    selected_spt = request.form.get("spt", "Semua")
    selected_kab = request.form.get("kab", "Semua")
    selected_kec = request.form.get("kec", "Semua")
    selected_cat = request.form.get("cat", "Semua")

    # ===== Filter bertingkat =====
    if selected_spt != "Semua":
        df_spt = df[df["observasi_no_spt"] == selected_spt]
    else:
        df_spt = df

    if selected_kab != "Semua":
        df_kab = df_spt[df_spt["observasi_kota_nama"] == selected_kab]
    else:
        df_kab = df_spt

    if selected_kec != "Semua":
        df_kec = df_kab[df_kab["observasi_kecamatan_nama"] == selected_kec]
    else:
        df_kec = df_kab

    # ===== Dropdown options =====
    spt_options = ["Semua"] + sorted(df["observasi_no_spt"].dropna().unique().tolist())
    kab_options = ["Semua"] + sorted(df_spt["observasi_kota_nama"].dropna().unique().tolist())
    kec_options = ["Semua"] + sorted(df_kab["observasi_kecamatan_nama"].dropna().unique().tolist())
    cat_options = ["Semua"] + sorted(df_kec["scan_catatan"].dropna().unique().tolist())

    # ===== Filter akhir untuk tampilan data =====
    filt = df.copy()
    if selected_spt != "Semua":
        filt = filt[filt["observasi_no_spt"] == selected_spt]
    if selected_kab != "Semua":
        filt = filt[filt["observasi_kota_nama"] == selected_kab]
    if selected_kec != "Semua":
        filt = filt[filt["observasi_kecamatan_nama"] == selected_kec]
    if selected_cat != "Semua":
        filt = filt[filt["scan_catatan"] == selected_cat]

    if filt.empty:
        return f"<h3>Data kosong untuk kombinasi tersebut.</h3><p>SPT: {selected_spt}, Kab: {selected_kab}, Kec: {selected_kec}</p>"
    
    # === Ringkasan Data untuk Info Cards ===
    total_data = len(filt)
    
    # Hitung jumlah ilegal (TANPA IZIN + KADALUARSA)
    filt_no_netral = filt[~filt["observasi_status_identifikasi_name"].isin(["CLEAR", "NOISE", "IDENTIFIKASI LEBIH LANJUT", "BELUM DIKETAHUI"])]
    total_no_netral = len(filt_no_netral)
    ilegal_count = filt[filt['observasi_status_identifikasi_name'].str.upper().isin(['TANPA IZIN', 'KADALUARSA'])].shape[0]
    ilegal_percent = round((ilegal_count / total_no_netral) * 100, 2) if total_no_netral else 0
    
    # Persentase berizin = 100 - ilegal
    berizin_percent = 100 - ilegal_percent

    # Hitung jumlah & persentase off air
    offair_count = filt[filt['observasi_status_identifikasi_name'].str.upper().str.contains('OFF AIR')].shape[0]
    offair_percent = round((offair_count / total_data) * 100, 1) if total_data else 0
    
    # Hitung teridentifikasi
    teridentifikasi_count = filt[filt['jenis'] == 'Teridentifikasi'].shape[0]
    
    # Hitung jumlah Kab/Kota termonitor
    jumlah_kota_termonitor = df['observasi_kota_nama'].nunique()
    
    # Hitung persentase kab/kota termonitor
    persen_kota_termonitor = int((jumlah_kota_termonitor / 10) * 100)

    
    try:
        # === Load ISR & Samakan Format ===
        base_dir = os.path.dirname(os.path.abspath(__file__))
        isr_path = os.path.join(base_dir, "Data Target Monitor ISR 2025 - Mataram.csv")
    
        # Load CSV
        df_ISR = pd.read_csv(isr_path, on_bad_lines='skip', delimiter=';') \
                    .rename(columns={'Freq': 'Frekuensi', 'Clnt Name': 'Identifikasi'})
        df_ISR['Frekuensi'] = pd.to_numeric(df_ISR['Frekuensi'], errors='coerce')
        filt['observasi_frekuensi'] = pd.to_numeric(filt['observasi_frekuensi'], errors='coerce')
    
        # Filter ISR bila kab dipilih
        if selected_kab != "Semua":
            df_ISR = df_ISR[df_ISR['Kab/Kota'].astype(str).str.strip().str.upper() == selected_kab.strip().upper()]
    
        # === Hitung kesesuaian dengan ISR ===
        freq_df1 = filt.groupby(
            ['observasi_frekuensi','observasi_sims_client_name','observasi_kota_nama']
        ).size().reset_index(name='Jumlah_df1')
    
        freq_df1 = freq_df1.rename(columns={
            'observasi_frekuensi': 'Frekuensi',
            'observasi_sims_client_name': 'Identifikasi',
            'observasi_kota_nama': 'Kab/Kota'
        })

        freq_df2 = df_ISR.groupby(['Frekuensi','Identifikasi','Kab/Kota']).size().reset_index(name='Jumlah_df2')
        
        # Hapus duplikat, hanya sisakan satu per kombinasi Frekuensi & Identifikasi
        merged = pd.merge(freq_df1, freq_df2, on=['Frekuensi', 'Identifikasi','Kab/Kota'], how='inner')
        
        jumlah_sesuai_isr = len(merged)


        # Ambil kota yang termonitor dari data observasi
        kota_termonitor = filt['observasi_kota_nama'].dropna().astype(str).str.strip().str.upper().unique()

        # Filter target ISR berdasarkan kota termonitor
        target_match = df_ISR[df_ISR['Kab/Kota'].str.strip().str.upper().isin(kota_termonitor)]
    
        # Hitung target ISR total (jumlah baris, karena tidak ada kolom "Jumlah ISR")
        jumlah_target_isr = len(target_match) if not target_match.empty else 0

        # Hitung persen kesesuaian
        persen_sesuai_isr = round((jumlah_sesuai_isr / jumlah_target_isr * 100), 2) if jumlah_target_isr > 0 else 0
    
    except Exception as e:
        print("Gagal hitung kesesuaian ISR:", e)
        jumlah_sesuai_isr, jumlah_target_isr, persen_sesuai_isr = 0, 0, 0


    # Persentase ISR sesuai target (sementara contoh statis)
    isr_percent = persen_sesuai_isr

    # Chart
    pie1 = px.pie(filt, names="observasi_status_identifikasi_name", title="Distribusi Legalitas",
                  hole=0.5, 
                  color_discrete_sequence=["#006db0", "#00ade6", "#edbc1b", "#8f181b", "#EF4444", "#6B7280",
                                           "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                           "#044065", "#d5ad2b", "#884a4c"])
    pie1.update_layout(
        legend=dict(
            font=dict(color='white', size=8))
        )
    pie_band = px.pie(filt, names="band_nama", title="Distribusi Band", 
                      hole=0.5,
                      color_discrete_sequence=["#006db0", "#00ade6", "#edbc1b", "#8f181b", "#EF4444", "#6B7280",
                                               "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                               "#044065", "#d5ad2b", "#884a4c"])

    bar = filt.groupby(["observasi_service_name", "jenis"]).size().reset_index(name="jumlah").sort_values(by="jumlah", ascending=False)
    total_per_dinas = (
    filt.groupby("observasi_service_name")
    .size()
    .sort_values(ascending=False))
    
    ordered_dinas = total_per_dinas.index.tolist()  # urutan berdasarkan total
    bar1 = px.bar(bar, y="observasi_service_name", x="jumlah", color="jenis", orientation='h',
                  title="Distribusi Dinas & Jenis",
                  labels={"observasi_service_name": "Nama Dinas",
                          "jumlah": "Jumlah Data",
                          "jenis": ""},
                  category_orders={"observasi_service_name": ordered_dinas},
                  color_discrete_sequence=["#006db0", "#00ade6", "#edbc1b", "#8f181b", "#EF4444", "#6B7280",
                                           "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                           "#044065", "#d5ad2b", "#884a4c"])
    
    bar1.update_layout(
        uniformtext_minsize=8,
        uniformtext_mode='hide',
        bargap=0.2,
        plot_bgcolor='#0f172a',
        paper_bgcolor='#0f172a',
        font=dict(color='white'),
        legend=dict(
            orientation="h",      # horizontal
            yanchor="bottom",     # anchor ke bawah
            y=-0.3,               # posisi agak di luar bawah chart
            xanchor="center",     
            x=0.5                 # posisi di tengah
        )
    )

    bar_pita = filt.groupby(["observasi_range_frekuensi", "observasi_status_identifikasi_name"]).size().reset_index(name="jumlah").sort_values(by="jumlah", ascending=False)
    # Buat kolom pita_singkat
    bar_pita["pita_singkat"] = bar_pita["observasi_range_frekuensi"].astype(str).str.split('.').str[0]

    total_per_pita = (
    filt.groupby("observasi_range_frekuensi")
    .size()
    .sort_values(ascending=False))
    
    ordered_pita = total_per_pita.index.tolist()  # urutan berdasarkan total
    # Urutan kategori singkat berdasarkan urutan original
    ordered_pita_singkat = [
        str(val).split('.')[0] for val in ordered_pita
    ]
    
    bar1_pita = px.bar(
        bar_pita,
        y="pita_singkat",
        x="jumlah",
        color="observasi_status_identifikasi_name",
        orientation='h',
        title="Distribusi Pita & Legalitas",
        labels={
            "pita_singkat": "Pita Frekuensi",
            "jumlah": "Jumlah Data",
            "observasi_status_identifikasi_name": ""
        },
        category_orders={"pita_singkat": ordered_pita_singkat},
        hover_name="observasi_range_frekuensi",
        color_discrete_sequence=["#006db0", "#00ade6", "#edbc1b", "#8f181b", "#EF4444", "#6B7280",
                                 "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                 "#044065", "#d5ad2b", "#884a4c"]
    )
    
    bar1_pita.update_traces(text=None)    
    bar1_pita.update_layout(
        uniformtext_minsize=8,
        uniformtext_mode='hide',
        bargap=0.2,
        plot_bgcolor='#0f172a',
        paper_bgcolor='#0f172a',
        legend=dict(
            font=dict(size=8),
            orientation="h",
            yanchor="bottom",
            y=-0.4,
            xanchor="center",
            x=0.5,
        )
    )
    
    pie_pantib = px.pie(
        df_pantib, 
        names="status_pelanggaran_name", 
        title="Jenis Pelanggaran",
        hole=0.5, 
        color_discrete_sequence=["#006db0", "#00ade6", "#edbc1b", "#8f181b", "#EF4444", "#6B7280",
                                 "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                 "#044065", "#d5ad2b", "#884a4c"])
    
    pie_pantib.update_layout(
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.05,
            font=dict(size=12)
        )
    )


    bar_data = df_pantib.groupby("status_pelanggaran_name").size().reset_index(name="jumlah")

    bar_pantib = px.bar(
        bar_data,
        x="status_pelanggaran_name",
        y="jumlah",
        title="Jumlah Pelanggaran",
        text="jumlah",
        labels={"status_pelanggaran_name":"Jenis Pelanggaran"},
        color_discrete_sequence=["#006db0", "#00ade6", "#edbc1b", "#8f181b", "#EF4444", "#6B7280",
                                 "#6d98b3", "#91cfe3", "#af8703", "#a83639", "#575759", "#252526",
                                 "#044065", "#d5ad2b", "#884a4c"]
    )

    for fig in [pie1]:
        for fig in [pie1, pie_band, bar1, bar1_pita, pie_pantib, bar_pantib, pie_inspeksi, bar_inspeksi]:
            fig.update_layout(
                paper_bgcolor="#1e293b",  # background luar chart
                plot_bgcolor="#1e293b",   # background area plot
                font=dict(color="white")  # teks jadi putih
            )
        
    for fig in [pie_band, bar1, bar1_pita, pie_pantib, bar_pantib, pie_inspeksi, bar_inspeksi]:
        fig.update_layout(
            paper_bgcolor="#1e293b",
            plot_bgcolor="#1e293b",
            font=dict(color="white"),
            margin=dict(l=40, r=20, t=60, b=80),
            legend=dict(
                orientation="h",     # horizontal
                yanchor="top",       # anchor ke atas legend
                y=-0.2,              # posisikan sedikit di bawah chart
                xanchor="center",
                x=0.5
            )
        )


    pie1_json = json.dumps(pie1, cls=plotly.utils.PlotlyJSONEncoder)
    pie_band_json = json.dumps(pie_band, cls=plotly.utils.PlotlyJSONEncoder)
    bar1_json = json.dumps(bar1, cls=plotly.utils.PlotlyJSONEncoder)
    bar1_pita_json = json.dumps(bar1_pita, cls=plotly.utils.PlotlyJSONEncoder)
    pie_pantib_json = json.dumps(pie_pantib, cls=PlotlyJSONEncoder)
    bar_pantib_json = json.dumps(bar_pantib, cls=PlotlyJSONEncoder)
    
    pie_inspeksi_json = json.dumps(pie_inspeksi, cls=PlotlyJSONEncoder)
    bar_inspeksi_json = json.dumps(bar_inspeksi, cls=PlotlyJSONEncoder)
    
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>DATA OBSERVASI BALAI MONITOR SFR KELAS II MATARAM TAHUN 2025</title>
        <link rel="icon" type="image/png" href="{{ url_for('static', filename='D.png') }}">
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            body { font-family: Arial; padding: 20px; }
            select { margin-right: 10px; }
            .chart { margin-top: 30px; }
            .chart-row {
                display: flex;
                flex-wrap: wrap;     /* biar chart bisa turun ke bawah kalau layar sempit */
                gap: 20px;
                margin: 20px 0;
            }
            }
            .chart {
                flex: 1;
                width: 100%;       /* ikut lebar container */
                height: auto;      /* tinggi menyesuaikan */
                min-width: 0;      /* biar bisa mengecil */
            }
            body {
                background-color: #0d1b2a;
                color: white;
                font-family: 'Segoe UI', sans-serif;
            }
        
            .filter-form {
                display: flex;
                flex-wrap: wrap;
                gap: 20px;
                margin: 25px 0;
                background-color: #1e1e1e; /* gelap */
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.3);
            }
            
            .filter-group label {
                font-weight: 600;
                margin-bottom: 5px;
                color: #ddd; /* teks terang */
            }
            
            .filter-group select {
                padding: 8px;
                border: 1px solid #444;
                border-radius: 6px;
                font-size: 14px;
                background-color: #2a2a2a; /* gelap */
                color: #f5f5f5; /* teks terang */
            }
            
            .filter-buttons button {
                padding: 10px 16px;
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 6px;
                font-weight: bold;
                cursor: pointer;
                transition: 0.3s;
            }
            
            .filter-buttons button:hover {
                background-color: #0056b3;
            }
            .chart-container {
                flex: 1 1 400px;     /* minimal 400px, tapi bisa melebar penuh */
                background-color: #1e293b;
                border-radius: 10px;
                padding: 15px;
                min-height: 400px;   /* tinggi minimal */
                width: 100%;         /* agar ikuti parent */
            }
        </style>
        
    </head>
    
    <script>
    function syncExcelForm() {
      document.getElementById('excel-spt').value = document.getElementById('spt').value;
      document.getElementById('excel-kab').value = document.getElementById('kab').value;
      document.getElementById('excel-kec').value = document.getElementById('kec').value;
      document.getElementById('excel-cat').value = document.getElementById('cat').value;
    }
    
    function autoSubmit(level) {
      const form = document.getElementById('main-form');
      const spt = document.getElementById('spt');
      const kab = document.getElementById('kab');
      const kec = document.getElementById('kec');
      const cat = document.getElementById('cat');
    
      // Reset bertingkat & kunci dropdown berikutnya
      if (level === 'spt') {
        kab.selectedIndex = 0; kab.disabled = false;
        kec.selectedIndex = 0; kec.disabled = true;
        cat.selectedIndex = 0; cat.disabled = true;
      } else if (level === 'kab') {
        kec.selectedIndex = 0; kec.disabled = false;
        cat.selectedIndex = 0; cat.disabled = true;
      } else if (level === 'kec') {
        cat.selectedIndex = 0; cat.disabled = false;
      }
    
      // Sinkron nilai ke form Excel
      syncExcelForm();
    
      // Auto-submit untuk refresh data di server
      // requestSubmit() akan hormati type & constraints, fallback ke submit()
      if (form.requestSubmit) form.requestSubmit();
      else form.submit();
    }
    
    // Saat halaman pertama kali dimuat, pastikan form Excel tersinkron
    document.addEventListener('DOMContentLoaded', syncExcelForm);
    </script>



    <body style="background-color:#0d1b2a; color:white; font-family:Segoe UI, sans-serif;">

    <div style="display:flex; align-items:center; gap:15px; padding:20px;">
        <img src="/static/logo-komdigi2.png" style="height:50px;">
        <img src="/static/djid.png" style="height:50px;">
        <div>
            <h2 style="margin:0;">Dashboard Observasi Frekuensi – Balmon SFR Kelas II Mataram</h2/>
            <p style="margin:0; font-size:16px; color:#9ca3af;">ROL Assistance Summary System (ROLASS)</p>
        </div>
    </div>
    
    <!-- Info Cards -->
    <div style="display:grid; grid-template-columns: repeat(5, 1fr); gap:15px; padding:20px;">
        
        <!-- Card 1 -->
        <div style="background:#1e293b; padding:8px; border-radius:8px; display:flex; align-items:center; gap:12px;">
            <img src="/static/1.png" alt="Data" style="width:36px; height:36px;">
            <div>
                <h1 style="margin:0;">{{ total_data }}</h1>
                <p style="margin:0;">Total Data Observasi</p>
            </div>
        </div>
        
        <!-- Card 2 -->
        <div style="background:#1e293b; padding:8px; border-radius:8px; display:flex; align-items:center; gap:12px;">
            <img src="/static/4.png" alt="Teridentifikasi" style="width:36px; height:36px;">
            <div>
                <h1 style="margin:0;">{{ teridentifikasi_count }}</h1>
                <p style="margin:0;">Teridentifikasi</p>
            </div>
        </div>
        
        <!-- Card 3 -->
        <div style="background:#1e293b; padding:8px; border-radius:8px; display:flex; align-items:center; gap:12px;">
            <img src="/static/2.png" alt="Berizin" style="width:36px; height:36px;">
            <div>
                <h1 style="margin:0;">{{ berizin_percent }}%</h1>
                <p style="margin:0;">Berizin</p>
            </div>
        </div>
        
        <!-- Card 4 -->
        <div style="background:#1e293b; padding:8px; border-radius:8px; display:flex; align-items:center; gap:12px;">
            <img src="/static/5.png" alt="ISR" style="width:36px; height:36px;">
            <div>
                <h1 style="margin:0;">{{ isr_percent }}%</h1>
                <p style="margin:0;">ISR Sesuai Target</p>
            </div>
        </div>
        
        <!-- Card 5 -->
        <div style="background:#1e293b; padding:8px; border-radius:8px; display:flex; align-items:center; gap:12px;">
            <img src="/static/6.png" alt="Kab/Kota Termonitor" style="width:36px; height:36px;">
            <div>
                <h1 style="margin:0;">{{ persen_kota_termonitor }}%</h1>
                <p style="margin:0;">Kab/Kota Termonitor</p>
            </div>
        </div>
    </div>

    
    <!-- Filter -->
    <form method="POST" class="filter-form" id="main-form">
        <!-- Dropdown filter -->
        <div class="filter-group">
            <label for="spt">No SPT</label>
            <select name="spt" id="spt" onchange="autoSubmit('spt')">
                {% for spt in spt_options %}
                <option value="{{ spt }}" {% if spt == selected_spt %}selected{% endif %}>{{ spt }}</option>
                {% endfor %}
            </select>
        </div>
    
        <!-- Tambahkan dropdown Kab/Kota -->
        <div class="filter-group">
            <label for="kab">Kab/Kota</label>
            <select name="kab" id="kab" onchange="autoSubmit('kab')">
                {% for kab in kab_options %}
                <option value="{{ kab }}" {% if kab == selected_kab %}selected{% endif %}>{{ kab }}</option>
                {% endfor %}
            </select>
        </div>
    
        <!-- Tambahkan dropdown Kecamatan -->
        <div class="filter-group">
            <label for="kec">Kecamatan</label>
            <select name="kec" id="kec" onchange="autoSubmit('kec')">
                {% for kec in kec_options %}
                <option value="{{ kec }}" {% if kec == selected_kec %}selected{% endif %}>{{ kec }}</option>
                {% endfor %}
            </select>
        </div>
    
        <!-- Tambahkan dropdown Catatan -->
        <div class="filter-group">
            <label for="cat">Catatan</label>
            <select name="cat" id="cat" onchange="autoSubmit('cat')">
                {% for cat in cat_options %}
                <option value="{{ cat }}" {% if cat == selected_cat %}selected{% endif %}>{{ cat }}</option>
                {% endfor %}
            </select>
        </div>
    
        <!-- Tombol -->
        <div style="display:flex; align-items:flex-end; gap:10px; padding:20px; justify-content:flex-end;">
            <button type="submit" style="background:#006db0; color:white; border:none; padding:8px 14px; border-radius:6px;">🔍 Tampilkan</button>
            <button form="excel-form" type="submit" style="background:#006db0; color:white; border:none; padding:8px 14px; border-radius:6px;">⬇️ Unduh Rekap</button>
            <button type="button" onclick="openModal()"
                    class="btn"
                    style="background:#006db0; color:white; padding:8px 14px; border-radius:6px; text-decoration:none;">
               📄 Unduh Laporan
            </button>
        </div>
    </form>
    


    
    <!-- Modal -->
    <div id="laporanModal"
         style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5);">
      <div style="background:white; width:400px; margin:100px auto; padding:20px; border-radius:8px; position:relative;">
        <h3>Isi Nama Pelaksana</h3>
    
        <!-- Form di dalam modal -->
        <form id="laporanForm" method="POST" action="/unduh_laporan">
          <!-- Hidden filter biar tetap terkirim -->
          <input type="hidden" name="spt" value="{{ selected_spt }}">
          <input type="hidden" name="kab" value="{{ selected_kab }}">
          <input type="hidden" name="kec" value="{{ selected_kec }}">
          <input type="hidden" name="cat" value="{{ selected_cat }}">
          
          <!-- Tanggal SPT -->
          <div style="margin-bottom:10px;">
            <label for="tgl_spt" style="font-weight:bold; display:block; margin-bottom:5px; color:#333;">
              📅 Tanggal SPT
            </label>
            <input type="date" name="tgl_spt" id="tgl_spt"
                   style="width:100%; padding:6px; border:1px solid #ccc; border-radius:5px;">
          </div>
          
          <!-- Pilihan Perangkat -->
            <div style="margin-bottom:10px;">
              <label for="perangkat" style="font-weight:bold; display:block; margin-bottom:5px; color:#333;">
                ⚙️ Perangkat SMFR
              </label>
              <select name="perangkat" id="perangkat"
                      style="width:100%; padding:6px; border:1px solid #ccc; border-radius:5px;">
                <option value="Tetap/Transportable TCI">Tetap/Transportable TCI</option>
                <option value="Tetap/Transportable LS Telcom">Tetap/Transportable LS Telcom</option>
                <option value="Bergerak R&S DDF205 Unit Mobil Isuzu Elf / Hilux Hitam">
                  Bergerak R&S DDF205 Unit Mobil Isuzu Elf / Hilux Hitam
                </option>
                <option value="Bergerak R&S DDF205 Unit Mobil Hilux Silver">
                  Bergerak R&S DDF205 Unit Mobil Hilux Silver
                </option>
                <option value="Jinjing R&S DDF007">Jinjing R&S DDF007</option>
                <option value="Jinjing R&S PR100">Jinjing R&S PR100</option>
              </select>
            </div>

          <!-- Nama Pelaksana -->
          <div id="pelaksana-container">
            <input type="text" name="pelaksana" placeholder="Nama Pelaksana"
                   style="width:100%; margin-bottom:10px; padding:6px;">
          </div>
    
          <!-- Tombol tambah pelaksana -->
          <button type="button" onclick="tambahPelaksana()"
                  style="background:#006db0; color:white; border:none; padding:6px 12px; border-radius:5px;">
            ➕ Tambah Pelaksana
          </button>
    
          <!-- Tombol aksi -->
          <div style="margin-top:15px; text-align:right;">
            <button type="button" onclick="closeModal()"
                    style="background:#ccc; border:none; padding:6px 12px; border-radius:5px;">
              Batal.
            </button>
            <button type="submit"
                    style="background:#006db0; color:white; border:none; padding:8px 14px; border-radius:6px;">
              ✅ Unduh
            </button>
          </div>
        </form>
      </div>
    </div>
    
    <script>
    function openModal() {
      document.getElementById("laporanModal").style.display = "block";
    }
    function closeModal() {
      document.getElementById("laporanModal").style.display = "none";
    }
    function tambahPelaksana() {
      const container = document.getElementById("pelaksana-container");
      const input = document.createElement("input");
      input.type = "text";
      input.name = "pelaksana";
      input.placeholder = "Nama Pelaksana";
      input.style = "width:100%; margin-bottom:10px; padding:6px;";
      container.appendChild(input);
    }
    </script>


    
    <!-- Form Unduh Rekap (disinkron otomatis oleh JS) -->
        <form method="POST" action="/download_excel" id="excel-form" style="display:none;">
          <input type="hidden" name="spt" id="excel-spt" value="{{ selected_spt }}">
          <input type="hidden" name="kab" id="excel-kab" value="{{ selected_kab }}">
          <input type="hidden" name="kec" id="excel-kec" value="{{ selected_kec }}">
          <input type="hidden" name="cat" id="excel-cat" value="{{ selected_cat }}">
        </form>
    
        <form method="POST" action="/unduh_laporan">
            <input type="hidden" name="spt" value="{{ selected_spt }}">
            <input type="hidden" name="kab" value="{{ selected_kab }}">
            <input type="hidden" name="kec" value="{{ selected_kec }}">
            <input type="hidden" name="cat" value="{{ selected_cat }}">
        </form>

    
    <!-- Charts -->
    <div class="chart-row">
        <div class="chart-container" id="pie1"></div>
        <div class="chart-container" id="bar1_pita"></div>      
    </div>
    
    <div class="chart-row">
        <div class="chart-container" id="pie_band"></div>
        <div class="chart-container" id="bar1"></div>
    </div>

    <!-- Info Cards Pantib -->
    <div style="display:flex; gap:15px; padding:20px; flex-wrap:wrap;">
        <div style="flex:1; min-width:200px; background:#1e293b; padding:15px; border-radius:8px; display:flex; align-items:center; gap:10px;">
            <div style="font-size:2rem;">⚠️</div>
            <div>
                <h1 style="margin:0;">{{ jumlah_pelanggaran }}</h1>
                <p>Jumlah Pelanggaran</p>
            </div>
        </div>
        <div style="flex:1; min-width:200px; background:#1e293b; padding:15px; border-radius:8px; display:flex; align-items:center; gap:10px;">
            <div style="font-size:2rem;">✅</div>
            <div>
                <h1 style="margin:0;">{{ persentase_ditertibkan }}%</h1>
                <p>Telah Ditertibkan</p>
            </div>
        </div>
    </div>

    <div class="chart-row">
        <div class="chart-container" id="pie_pantib"></div>
        <div class="chart-container" id="bar_pantib"></div>
    </div>
        
    <!-- Info Cards Inspeksi -->
    <div style="display:flex; gap:15px; padding:20px; flex-wrap:wrap;">
        <div style="flex:1; min-width:200px; background:#1e293b; padding:15px; border-radius:8px; display:flex; align-items:center; gap:10px;">
            <div style="font-size:2rem;">📊</div>
            <div>
                <h1 style="margin:0;">{{ total_inspeksi }}</h1>
                <p>Total Inspeksi</p>
            </div>
        </div>
        <div style="flex:1; min-width:200px; background:#1e293b; padding:15px; border-radius:8px; display:flex; align-items:center; gap:10px;">
            <div style="font-size:2rem;">🔍</div>
            <div>
                <h1 style="margin:0;">{{ sudah_inspeksi }}</h1>
                <p>Sudah Inspeksi</p>
            </div>
        </div>
    </div>

    <div class="chart-row">
        <div class="chart-container" id="pie_inspeksi"></div>
        <div class="chart-container" id="bar_inspeksi"></div>
    </div>

    <script>
        Plotly.newPlot("pie1", {{ pie1_json|safe }}.data, {{ pie1_json|safe }}.layout, {responsive: true});
        Plotly.newPlot("bar1_pita", {{ bar1_pita_json|safe }}.data, {{ bar1_pita_json|safe }}.layout, {responsive: true});
        Plotly.newPlot("pie_band", {{ pie_band_json|safe }}.data, {{ pie_band_json|safe }}.layout, {responsive: true});
        Plotly.newPlot("bar1", {{ bar1_json|safe }}.data, {{ bar1_json|safe }}.layout, {responsive: true});
    </script>
    
    <script>
        var piePantib = {{ pie_pantib_json|safe }};
        var barPantib = {{ bar_pantib_json|safe }};
        Plotly.newPlot('pie_pantib', piePantib.data, piePantib.layout, {responsive:true});
        Plotly.newPlot('bar_pantib', barPantib.data, barPantib.layout, {responsive:true});
    </script>
    
    <script>
        var pieInspeksi = {{ pie_inspeksi_json|safe }};
        var barInspeksi = {{ bar_inspeksi_json|safe }};
        Plotly.newPlot('pie_inspeksi', pieInspeksi.data, pieInspeksi.layout, {responsive:true});
        Plotly.newPlot('bar_inspeksi', barInspeksi.data, barInspeksi.layout, {responsive:true});
    </script>
    
    </body>
    </html>
    ''',
    spt_options=spt_options,
    kab_options=kab_options,
    kec_options=kec_options,
    cat_options=cat_options,
    selected_spt=selected_spt,
    selected_kab=selected_kab,
    selected_kec=selected_kec,
    selected_cat=selected_cat,
    pie1_json=pie1_json,
    pie_band_json=pie_band_json,
    bar1_json=bar1_json,
    bar1_pita_json=bar1_pita_json,
    total_data=total_data,
    berizin_percent=berizin_percent,
    offair_percent=offair_percent,
    persen_kota_termonitor = persen_kota_termonitor,
    teridentifikasi_count=teridentifikasi_count,
    isr_percent=isr_percent,
    jumlah_pelanggaran=jumlah_pelanggaran,
    persentase_ditertibkan=persentase_ditertibkan,
    pie_pantib=pie_pantib.to_html(full_html=False),
    bar_pantib=bar_pantib.to_html(full_html=False),
    pie_pantib_json=pie_pantib_json,
    bar_pantib_json=bar_pantib_json,
    sudah_inspeksi = sudah_inspeksi,
    total_inspeksi = total_inspeksi,
    capaian_inspeksi = capaian_inspeksi,
    pie_inspeksi_json=pie_inspeksi_json,
    bar_inspeksi_json=bar_inspeksi_json
    )
    
@app.route("/get_kab/<spt>")
def get_kab(spt):
    df = load_data()
    if spt != "Semua":
        df = df[df["observasi_no_spt"] == spt]
    kab_options = sorted(df["observasi_kota_nama"].dropna().unique().tolist())
    return {"kab_list": kab_options}

@app.route("/get_kec/<spt>/<kab>")
def get_kec(spt, kab):
    df = load_data()
    if spt != "Semua":
        df = df[df["observasi_no_spt"] == spt]
    if kab != "Semua":
        df = df[df["observasi_kota_nama"] == kab]
    kec_options = sorted(df["observasi_kecamatan_nama"].dropna().unique().tolist())
    return {"kec_list": kec_options}

@app.route("/get_cat/<spt>/<kab>/<kec>")
def get_cat(spt, kab, kec):
    df = load_data()
    if spt != "Semua":
        df = df[df["observasi_no_spt"] == spt]
    if kab != "Semua":
        df = df[df["observasi_kota_nama"] == kab]
    if kec != "Semua":
        df = df[df["observasi_kecamatan_nama"] == kec]
    cat_options = sorted(df["scan_catatan"].dropna().unique().tolist())
    return {"cat_list": cat_options}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=80)














