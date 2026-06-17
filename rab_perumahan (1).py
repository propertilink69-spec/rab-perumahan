import pandas as pd
import os
from datetime import datetime
import json
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import locale
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import sys

# Set locale
try:
    locale.setlocale(locale.LC_ALL, 'id_ID.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, '')

# File data
RAB_FILE = 'rab_perumahan.xlsx'
MATERIALS_FILE = 'kebutuhan_bahan.xlsx'
DAILY_EXP_FILE = 'laporan_belanja_harian.xlsx'
USAGE_FILE = 'penggunaan_material.xlsx'
PROJECT_INFO = 'project_info.json'
CREDENTIALS_FILE = 'credentials.json'  # Google Service Account

class RABApp:
    def __init__(self):
        self.load_project_info()
        self.init_files()
        self.gc = None  # Google Sheets client
        self.drive_service = None

    def load_project_info(self):
        if os.path.exists(PROJECT_INFO):
            with open(PROJECT_INFO, 'r') as f:
                self.project = json.load(f)
        else:
            self.project = {
                "nama_proyek": "Pembangunan Perumahan XYZ",
                "lokasi": "Jakarta",
                "luas_tanah": 500,
                "jumlah_unit": 10,
                "tanggal_mulai": datetime.now().strftime("%Y-%m-%d"),
                "nama_perusahaan": "PT. Kontraktor Sejahtera",
                "logo_path": "logo_perusahaan.png",
                "ttd_path": "ttd_kepala_proyek.png",
                "nama_penanda": "Ir. Ahmad Santoso",
                "jabatan": "Kepala Proyek",
                "google_sheet_id": ""  # Untuk integrasi Google Sheets
            }
            self.save_project_info()

    def save_project_info(self):
        with open(PROJECT_INFO, 'w') as f:
            json.dump(self.project, f, indent=4)

    def init_google_sheets(self):
        if not os.path.exists(CREDENTIALS_FILE):
            print("⚠️ File credentials.json belum ada.")
            print("Silakan buat Service Account di Google Cloud dan simpan sebagai credentials.json")
            return False
        try:
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
            self.gc = gspread.authorize(creds)
            self.drive_service = build('drive', 'v3', credentials=creds)
            print("✅ Terhubung ke Google Sheets & Drive!")
            return True
        except Exception as e:
            print(f"❌ Gagal koneksi: {e}")
            return False

    def upload_to_drive(self, file_path, folder_name="RAB_Proyek"):
        """Upload file ke Google Drive"""
        if not self.drive_service:
            if not self.init_google_sheets():
                return False
        try:
            # Cari atau buat folder
            folder_query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'"
            results = self.drive_service.files().list(q=folder_query, fields="files(id)").execute()
            folders = results.get('files', [])
            
            if folders:
                folder_id = folders[0]['id']
            else:
                file_metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
                folder = self.drive_service.files().create(body=file_metadata, fields='id').execute()
                folder_id = folder.get('id')
            
            # Upload file
            file_metadata = {'name': os.path.basename(file_path), 'parents': [folder_id]}
            media = MediaFileUpload(file_path, resumable=True)
            file = self.drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            
            print(f"✅ Berhasil upload ke Google Drive: {os.path.basename(file_path)}")
            print(f"Folder: {folder_name}")
            return True
        except Exception as e:
            print(f"❌ Gagal upload ke Drive: {e}")
            return False

    def sync_to_sheets(self, sheet_name, df, worksheet_name="Data"):
        if not self.gc:
            if not self.init_google_sheets():
                return False
        try:
            spreadsheet = self.gc.open_by_key(self.project.get("google_sheet_id")) if self.project.get("google_sheet_id") else self.gc.create(sheet_name)
            if not self.project.get("google_sheet_id"):
                self.project["google_sheet_id"] = spreadsheet.id
                self.save_project_info()
            
            try:
                worksheet = spreadsheet.worksheet(worksheet_name)
            except:
                worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=20)
            
            worksheet.clear()
            worksheet.update([df.columns.values.tolist()] + df.values.tolist())
            print(f"✅ Data berhasil disinkron ke Google Sheets: {worksheet_name}")
            return True
        except Exception as e:
            print(f"❌ Gagal sync ke Sheets: {e}")
            return False

    # ... (semua method lama tetap sama, tambahkan menu sync)

    def init_files(self):
        # RAB
        if not os.path.exists(RAB_FILE):
            data_rab = {
                'Kategori': ['Persiapan', 'Fondasi', 'Struktur', 'Dinding', 'Atap', 'Instalasi', 'Finishing', 'Lain-lain'],
                'Uraian Pekerjaan': ['Pembersihan lahan', 'Pengecoran fondasi', 'Kolom & Balok', 'Pasang bata', 'Pasang genteng', 'Listrik & Plumbing', 'Cat & Keramik', ''],
                'Volume': [1, 50, 100, 2000, 500, 1, 1, 1],
                'Satuan': ['Lot', 'm3', 'm3', 'buah', 'm2', 'Lot', 'Lot', 'Lot'],
                'Harga Satuan (Rp)': [5000000, 1200000, 1500000, 800, 45000, 25000000, 15000000, 5000000],
                'Total Biaya (Rp)': [0]*8
            }
            df = pd.DataFrame(data_rab)
            df['Total Biaya (Rp)'] = df['Volume'] * df['Harga Satuan (Rp)']
            df.to_excel(RAB_FILE, index=False)

        if not os.path.exists(MATERIALS_FILE):
            data_mat = {
                'Bahan': ['Semen', 'Pasir', 'Batu Split', 'Besi Beton', 'Bata Merah', 'Keramik', 'Cat', 'Genteng'],
                'Spesifikasi': ['Portland', 'Pasir Beton', 'Split 1-2 cm', 'Ø10 & Ø12', 'Merah Press', '60x60', 'Tembok Luar', 'Genteng Metal'],
                'Stok Awal': [0, 0, 0, 0, 0, 0, 0, 0],
                'Jumlah Dibutuhkan': [500, 200, 150, 10000, 15000, 800, 50, 1000],
                'Satuan': ['sak', 'm3', 'm3', 'kg', 'buah', 'm2', 'kaleng', 'buah'],
                'Harga Satuan': [65000, 250000, 300000, 18000, 800, 120000, 45000, 25000],
                'Total Estimasi': [0]*8,
                'Stok Saat Ini': [0]*8
            }
            df = pd.DataFrame(data_mat)
            df['Total Estimasi'] = df['Jumlah Dibutuhkan'] * df['Harga Satuan']
            df['Stok Saat Ini'] = df['Stok Awal']
            df.to_excel(MATERIALS_FILE, index=False)

        if not os.path.exists(DAILY_EXP_FILE):
            pd.DataFrame(columns=['Tanggal', 'Uraian', 'Kategori', 'Jumlah (Rp)', 'Supplier', 'Keterangan']).to_excel(DAILY_EXP_FILE, index=False)

        if not os.path.exists(USAGE_FILE):
            pd.DataFrame(columns=['Tanggal', 'Bahan', 'Jumlah Digunakan', 'Satuan', 'Keterangan']).to_excel(USAGE_FILE, index=False)

    def update_stok(self, bahan, jumlah, jenis='tambah'):
        df = pd.read_excel(MATERIALS_FILE)
        idx = df[df['Bahan'] == bahan].index
        if not idx.empty:
            if jenis == 'tambah':
                df.loc[idx, 'Stok Saat Ini'] += jumlah
            else:
                df.loc[idx, 'Stok Saat Ini'] -= jumlah
                if df.loc[idx, 'Stok Saat Ini'].values[0] < 0:
                    df.loc[idx, 'Stok Saat Ini'] = 0
            df.to_excel(MATERIALS_FILE, index=False)
            return True
        return False

    def cek_stok_rendah(self):
        df = pd.read_excel(MATERIALS_FILE)
        rendah = df[df['Stok Saat Ini'] < (df['Jumlah Dibutuhkan'] * 0.2)]
        if not rendah.empty:
            print("\n⚠️  PERINGATAN STOK RENDAH:")
            print(rendah[['Bahan', 'Stok Saat Ini', 'Jumlah Dibutuhkan']].to_string(index=False))

    def format_rupiah(self, amount):
        return f"Rp {amount:,.0f}"

    def add_logo_to_pdf(self, elements):
        logo_path = self.project.get("logo_path", "logo_perusahaan.png")
        if os.path.exists(logo_path):
            try:
                logo = Image(logo_path, width=150, height=70)
                elements.append(logo)
                elements.append(Spacer(1, 12))
            except:
                pass

    def add_signature_to_pdf(self, elements):
        ttd_path = self.project.get("ttd_path", "ttd_kepala_proyek.png")
        if os.path.exists(ttd_path):
            try:
                elements.append(Spacer(1, 50))
                styles = getSampleStyleSheet()
                sig_data = [[
                    Paragraph(f"<b>{self.project.get('nama_penanda', '')}</b><br/>{self.project.get('jabatan', '')}<br/>Tanggal: {datetime.now().strftime('%d %B %Y')}", styles['Normal']),
                    Image(ttd_path, width=180, height=90)
                ]]
                sig_table = Table(sig_data, colWidths=[280, 200])
                sig_table.setStyle(TableStyle([('ALIGN', (1,0), (1,0), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
                elements.append(sig_table)
            except:
                pass

    def export_to_pdf(self, report_type):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        
        if report_type == 'rab':
            df = pd.read_excel(RAB_FILE)
            filename = f'RAB_{self.project["nama_proyek"].replace(" ", "_")}_{timestamp}.pdf'
            title = "RENCANA ANGGARAN BIAYA (RAB)"
            total_col = 'Total Biaya (Rp)'
            total = df[total_col].sum() if not df.empty else 0
        elif report_type == 'stok':
            df = pd.read_excel(MATERIALS_FILE)
            filename = f'Stok_Material_{timestamp}.pdf'
            title = "LAPORAN STOK MATERIAL"
            total_col = 'Total Estimasi'
            total = df[total_col].sum() if total_col in df.columns else 0
        elif report_type == 'belanja':
            df = pd.read_excel(DAILY_EXP_FILE)
            filename = f'Laporan_Belanja_Harian_{timestamp}.pdf'
            title = "LAPORAN BELANJA HARIAN"
            total_col = 'Jumlah (Rp)'
            total = df[total_col].sum() if not df.empty else 0
        elif report_type == 'penggunaan':
            df = pd.read_excel(USAGE_FILE)
            filename = f'Penggunaan_Material_{timestamp}.pdf'
            title = "LAPORAN PENGGUNAAN MATERIAL"
            total = 0
        else:
            return

        doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=80, bottomMargin=60)
        elements = []
        styles = getSampleStyleSheet()

        self.add_logo_to_pdf(elements)

        elements.append(Paragraph(f"<b>{self.project.get('nama_perusahaan', '')}</b>", styles['Heading1']))
        elements.append(Paragraph(f"<b>{title}</b>", styles['Title']))
        elements.append(Paragraph(f"Proyek: {self.project['nama_proyek']}", styles['Normal']))
        elements.append(Paragraph(f"Lokasi: {self.project['lokasi']}", styles['Normal']))
        elements.append(Paragraph(f"Tanggal Cetak: {datetime.now().strftime('%d %B %Y %H:%M')}", styles['Normal']))
        elements.append(Spacer(1, 20))

        data = [df.columns.tolist()] + df.values.tolist()
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
        ]))

        elements.append(table)
        if total > 0:
            elements.append(Spacer(1, 20))
            elements.append(Paragraph(f"<b>Total: {self.format_rupiah(total)}</b>", styles['Heading2']))

        self.add_signature_to_pdf(elements)
        doc.build(elements)
        print(f"✅ PDF diekspor: {filename}")

        # Otomatis upload ke Google Drive setelah ekspor
        if input("Upload ke Google Drive? (y/n): ").lower() == 'y':
            self.upload_to_drive(filename)

    # Method lama lainnya (tambah_belanja, catat_penggunaan, dll) tetap
    def tambah_belanja_harian(self):
        df = pd.read_excel(DAILY_EXP_FILE)
        tanggal = datetime.now().strftime("%Y-%m-%d")
        uraian = input("Uraian belanja: ")
        kategori = input("Kategori (Material/Upah/Lainnya): ")
        jumlah_rp = float(input("Jumlah (Rp): "))
        supplier = input("Supplier: ")
        keterangan = input("Keterangan: ")

        new_row = pd.DataFrame([{'Tanggal': tanggal, 'Uraian': uraian, 'Kategori': kategori,
                                'Jumlah (Rp)': jumlah_rp, 'Supplier': supplier, 'Keterangan': keterangan}])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(DAILY_EXP_FILE, index=False)

        if kategori.lower() == 'material':
            mat_df = pd.read_excel(MATERIALS_FILE)
            print(mat_df[['Bahan', 'Satuan']])
            bahan = input("Nama bahan: ")
            jumlah = float(input("Jumlah dibeli: "))
            self.update_stok(bahan, jumlah, 'tambah')

        print("✅ Belanja tercatat!")

    def catat_penggunaan_material(self):
        df = pd.read_excel(USAGE_FILE)
        tanggal = datetime.now().strftime("%Y-%m-%d")
        mat_df = pd.read_excel(MATERIALS_FILE)
        print(mat_df[['Bahan', 'Stok Saat Ini']])
        bahan = input("Bahan digunakan: ")
        jumlah = float(input("Jumlah: "))
        keterangan = input("Keterangan: ")

        satuan = mat_df[mat_df['Bahan']==bahan]['Satuan'].values[0] if not mat_df[mat_df['Bahan']==bahan].empty else ''
        new_row = pd.DataFrame([{'Tanggal': tanggal, 'Bahan': bahan, 'Jumlah Digunakan': jumlah, 'Satuan': satuan, 'Keterangan': keterangan}])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(USAGE_FILE, index=False)
        self.update_stok(bahan, jumlah, 'kurang')
        self.cek_stok_rendah()

    def tampilkan_stok(self):
        df = pd.read_excel(MATERIALS_FILE)
        print("\n=== STOK MATERIAL ===")
        print(df.to_string(index=False))
        self.cek_stok_rendah()

    def tampilkan_rab(self):
        df = pd.read_excel(RAB_FILE)
        total = df['Total Biaya (Rp)'].sum()
        print("\n=== RAB ===")
        print(df.to_string(index=False))
        print(f"Total: {self.format_rupiah(total)}")

    def update_rab(self):
        df = pd.read_excel(RAB_FILE)
        print(df[['Kategori', 'Uraian Pekerjaan', 'Volume', 'Harga Satuan (Rp)']])
        idx = int(input("Baris (0-based): "))
        if 0 <= idx < len(df):
            df.loc[idx, 'Volume'] = float(input("Volume baru: "))
            df.loc[idx, 'Harga Satuan (Rp)'] = float(input("Harga baru: "))
            df['Total Biaya (Rp)'] = df['Volume'] * df['Harga Satuan (Rp)']
            df.to_excel(RAB_FILE, index=False)
            print("✅ Diupdate!")

    def laporan_harian(self):
        df = pd.read_excel(DAILY_EXP_FILE)
        print("\n=== LAPORAN HARIAN ===")
        print(df.to_string(index=False))
        if not df.empty:
            print(f"Total: {self.format_rupiah(df['Jumlah (Rp)'].sum())}")

    def menu(self):
        while True:
            print("\n" + "="*80)
            print("   APLIKASI RAB + GOOGLE SHEETS INTEGRATION")
            print("="*80)
            print("1. Tampilkan RAB")
            print("2. Update RAB")
            print("3. Tambah Belanja Harian")
            print("4. Catat Penggunaan Material")
            print("5. Lihat Stok Material")
            print("6. Laporan Belanja Harian")
            print("7. Ekspor ke PDF (+ Drive)")
            print("8. Sync ke Google Sheets")
            print("9. Upload File ke Google Drive")
            print("10. Setting Proyek / Google")
            print("0. Keluar")
            pilihan = input("\nPilih: ")

            if pilihan == '1': self.tampilkan_rab()
            elif pilihan == '2': self.update_rab()
            elif pilihan == '3': self.tambah_belanja_harian()
            elif pilihan == '4': self.catat_penggunaan_material()
            elif pilihan == '5': self.tampilkan_stok()
            elif pilihan == '6': self.laporan_harian()
            elif pilihan == '7':
                print("1.RAB 2.Stok 3.Belanja 4.Penggunaan")
                sub = input("Pilih: ")
                if sub == '1': self.export_to_pdf('rab')
                elif sub == '2': self.export_to_pdf('stok')
                elif sub == '3': self.export_to_pdf('belanja')
                elif sub == '4': self.export_to_pdf('penggunaan')
            elif pilihan == '8':
                print("\nSync apa?")
                print("1. RAB 2. Stok 3. Belanja 4. Penggunaan 5. Semua")
                sub = input("Pilih: ")
                if sub == '1': self.sync_to_sheets("RAB_Proyek", pd.read_excel(RAB_FILE), "RAB")
                elif sub == '2': self.sync_to_sheets("Stok_Material", pd.read_excel(MATERIALS_FILE), "Stok")
                elif sub == '3': self.sync_to_sheets("Belanja_Harian", pd.read_excel(DAILY_EXP_FILE), "Belanja")
                elif sub == '4': self.sync_to_sheets("Penggunaan_Material", pd.read_excel(USAGE_FILE), "Penggunaan")
                elif sub == '5':
                    self.sync_to_sheets("RAB_Proyek", pd.read_excel(RAB_FILE), "RAB")
                    self.sync_to_sheets("Stok_Material", pd.read_excel(MATERIALS_FILE), "Stok")
            elif pilihan == '9':
                file_to_upload = input("Masukkan nama file untuk diupload (contoh: RAB_xxx.pdf): ")
                if os.path.exists(file_to_upload):
                    self.upload_to_drive(file_to_upload)
                else:
                    print("❌ File tidak ditemukan!")
            elif pilihan == '10':
                print(json.dumps(self.project, indent=2))
                if input("Ubah Google Sheet ID? (y/n): ").lower() == 'y':
                    self.project["google_sheet_id"] = input("Masukkan Google Sheet ID (dari URL): ")
                    self.save_project_info()
                if input("Ubah logo/TTD? (y/n): ").lower() == 'y':
                    self.project["logo_path"] = input("Logo path: ") or self.project.get("logo_path")
                    self.save_project_info()
            elif pilihan == '0':
                print("Terima kasih!")
                break

if __name__ == "__main__":
    try:
        import reportlab
    except:
        os.system("pip install reportlab")
    app = RABApp()
    app.menu()
