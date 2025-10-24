import os
import datetime
import pandas as pd
from flask import Flask, render_template, request, jsonify

# --- Pustaka GSheets & Localization ---
import gspread
from gspread_dataframe import set_with_dataframe
import json
import locale

# --- Konfigurasi Lokal dan App ---
# Atur locale ke Bahasa Indonesia (Penting untuk Nama Hari)
try:
    locale.setlocale(locale.LC_TIME, 'id_ID.utf8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'id_ID')
    except locale.Error:
        # Fallback jika locale Indonesia tidak tersedia (nama hari akan bahasa Inggris)
        pass 

app = Flask(__name__)
SHEET_TITLE = "Rekap Absensi Tapak Suci Sidayu" # Ganti dengan nama Google Sheet Anda


# --- FUNGSI KONEKSI DAN MANAJEMEN GSHEETS ---

def get_gspread_client():
    """Menginisialisasi klien GSpread menggunakan kredensial dari Environment Variable."""
    try:
        creds_json = os.environ.get('GSPREAD_CREDENTIALS')
        if not creds_json:
            print("ERROR: Environment variable GSPREAD_CREDENTIALS tidak ditemukan!")
            return None
        
        creds = json.loads(creds_json)
        gc = gspread.service_account_from_dict(creds)
        return gc
    except Exception as e:
        print(f"Gagal menginisialisasi GSpread client: {e}")
        return None

def get_weekly_worksheet_name():
    """Mengembalikan nama sheet dengan format HARI_TANGGAL (contoh: JUMAT_25-10-2025)."""
    today = datetime.datetime.now()
    
    # %A (Nama Hari Penuh Lokal), %d-%m-%Y (Tanggal-Bulan-Tahun)
    formatted_string = today.strftime("%A, %d-%m-%Y")
    
    # Ubah menjadi huruf kapital dan ganti koma dengan underscore
    return formatted_string.upper().replace(', ', '_')

def get_attendance_dataframe(gc):
    """Membaca data absensi dari sheet mingguan saat ini atau membuat sheet baru."""
    current_worksheet_name = get_weekly_worksheet_name() 
    columns = ["Kode", "Nama", "Waktu"]

    if not gc:
        return pd.DataFrame(columns=columns), "Gagal koneksi GSheets"

    try:
        sh = gc.open(SHEET_TITLE)
    except gspread.WorksheetNotFound:
        print(f"Sheet '{SHEET_TITLE}' tidak ditemukan!")
        return pd.DataFrame(columns=columns), "Sheet Gagal"

    try:
        # Mencoba membuka sheet mingguan saat ini
        ws = sh.worksheet(current_worksheet_name)
    except gspread.WorksheetNotFound:
        # Jika sheet minggu ini belum ada, buat yang baru
        ws = sh.add_worksheet(title=current_worksheet_name, rows="100", cols="3")
        
        # Inisialisasi header untuk sheet baru
        ws.append_row(columns)
        print(f"Worksheet baru '{current_worksheet_name}' telah dibuat.")
        return pd.DataFrame(columns=columns), sh # Kembalikan DataFrame kosong

    # Baca data dari worksheet ke Pandas DataFrame
    data = ws.get_all_values()
    
    # Jika data hanya header atau kosong
    if len(data) <= 1:
        return pd.DataFrame(columns=columns), sh
    
    # Konversi data ke DataFrame (baris pertama adalah header)
    df = pd.DataFrame(data[1:], columns=data[0])
    return df, sh

def save_attendance_dataframe(df, sh):
    """Menulis seluruh DataFrame absensi kembali ke Google Sheet."""
    current_worksheet_name = get_weekly_worksheet_name()
    try:
        ws = sh.worksheet(current_worksheet_name)
        
        # Tulis DataFrame ke Google Sheet (termasuk header)
        set_with_dataframe(ws, df, include_index=False, row=1, col=1)
        print(f"Data absensi berhasil disimpan ke Google Sheet: {SHEET_TITLE}/{current_worksheet_name}")
    except Exception as e:
        print(f"Gagal menyimpan data ke GSheets: {e}")


# --- FUNGSI INTI ABSENSI ---

def process_qrcode(qr_data, gc):
    df_absensi, sh = get_attendance_dataframe(gc)
    time_now_hms = datetime.datetime.now().strftime("%H:%M:%S") 
    
    if df_absensi.empty and isinstance(sh, str): # sh berupa string error
         return {
            "nama": "Gagal Koneksi GSheets",
            "status": "CONNECTION_ERROR",
            "waktu": time_now_hms
        }

    # 1. Parsing data QR: KODE_NAMA-LENGKAP-SISWA
    try:
        if '_' not in qr_data:
            raise ValueError()
            
        kode, nama_raw = qr_data.split('_', 1)
        nama_lengkap = nama_raw.replace('-', ' ').upper() 
        kode_upper = kode.upper()
        
    except ValueError:
        return {
            "nama": "Format QR Tidak Valid",
            "status": "INVALID_FORMAT",
            "waktu": time_now_hms
        }

    # 2. Cek apakah Kode sudah terdaftar di sheet hari ini
    if kode_upper in df_absensi['Kode'].values:
        waktu_sebelumnya = df_absensi[df_absensi['Kode'] == kode_upper]['Waktu'].iloc[0]
        result = {
            "nama": nama_lengkap,
            "kode": kode_upper,
            "status": "REGISTERED",
            "waktu": waktu_sebelumnya
        }
    else:
        # 3. Tambahkan data baru
        new_row = {"Kode": kode_upper, "Nama": nama_lengkap, "Waktu": time_now_hms}
        # Gunakan loc untuk menambah baris baru
        df_absensi.loc[len(df_absensi)] = new_row
        save_attendance_dataframe(df_absensi, sh) # Simpan ke GSheets
        
        result = {
            "nama": nama_lengkap,
            "kode": kode_upper,
            "status": "SUCCESS",
            "waktu": time_now_hms
        }

    # 4. Mengambil data absensi terbaru untuk ditampilkan
    recent_attendance = df_absensi.sort_values(by='Waktu', ascending=False).head(10)[['Nama', 'Waktu']].to_dict('records')
    result['recent_attendance'] = recent_attendance
    
    return result


# --- ROUTING FLASK ---

# Inisialisasi gspread client sekali saat server mulai
GSHEET_CLIENT = get_gspread_client()

@app.route('/')
def index():
    df_absensi, _ = get_attendance_dataframe(GSHEET_CLIENT)
    
    # Ambil 10 data terbaru untuk ditampilkan di HTML
    if df_absensi.empty:
        recent_attendance = []
    else:
        recent_attendance = df_absensi.sort_values(by='Waktu', ascending=False).head(10)[['Nama', 'Waktu']].to_dict('records')
        
    return render_template('index.html', recent_attendance=recent_attendance)

@app.route('/scan', methods=['POST'])
def scan():
    qr_data = request.json.get('qr_data')
    if not qr_data:
        return jsonify({"error": "No QR data provided"}), 400
    
    result = process_qrcode(qr_data, GSHEET_CLIENT)
    return jsonify(result)

if __name__ == '__main__':
    # Gunakan PORT dari Render Environment Variable, dan hapus SSL context
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)