import fitz  # PyMuPDF
import re
import pandas as pd
import os
import io
import json
import base64
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

def parse_amount(amount_str):
    """Mengonversi string angka gaya Eropa (mis., '1.234,56') menjadi float."""
    if not isinstance(amount_str, str):
        return 0.0
    # Menghapus titik ribuan dan mengganti koma desimal dengan titik
    cleaned = amount_str.replace('.', '').replace(',', '.')
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return 0.0

def process_pdf_to_excel(pdf_content, filename):
    """Fungsi utama untuk mengekstrak data dari konten PDF dan menghasilkan file Excel."""
    doc = fitz.open(stream=pdf_content, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()

    # Fungsi bantu untuk mencari pola dengan regex, mengembalikan default jika tidak ditemukan
    def find(pattern, default='N/A', source_text=text):
        match = re.search(pattern, source_text, re.DOTALL)
        return match.group(1).strip() if match else default

    # --- Ekstraksi Data Header ---
    invoice_number = find(r'Rechnungs Nr\.:\s*(\d+)')
    sender = find(r'Absender:\s*([^\n]+)')
    etd_eta = find(r'ETD/ETA:\s*([^\n]+)')
    port_loading = find(r'Port of Loading:\s*([^\n]+)')
    port_discharge = find(r'Port of Discharge:\s*([^\n]+)')
    invoice_date = find(r'Rechnungsdatum:\s*(\d{2}-[A-Za-z]{3}-\d{4})')
    stt_number = find(r'STT Nr\.:\s*(\d+)')
    
    # Ekstraksi berat dan volume dengan lebih spesifik
    gross_weight_kg = find(r'Bruttogewicht\s*([\d.,]+)\s*KGS')
    volume_cbm = find(r'Volumen\s*([\d.,]+)\s*CBM')

    # --- Ekstraksi Rincian Biaya ---
    # Memetakan nama biaya di PDF ke kode yang diinginkan
    cost_label_map = {
        "Summarische Eingangsmeldung": "ENS",
        "Seefracht": "SFRT",
        "THC (Terminal Handling Charge)": "THC",
        "Abfertigungskosten im": "CCDE",
        "ISPS (Hafen & Terminal": "ISPS",
        "Nachlaufkosten": "NL",
        "Delivery-/Drop-Off-Gebühr": "DROP",
        "Importverzollung in NL": "Zoll"
    }
    
    # Membatasi pencarian ke bagian "Unsere Leistungen" untuk akurasi
    cost_section = find(r"Unsere Leistungen(.*?)Gesamtkosten", default="")
    rows = []

    for label_in_pdf, code in cost_label_map.items():
        # Pola regex dinamis untuk setiap jenis biaya
        cost_pattern = rf"{re.escape(label_in_pdf)}.*?EUR\s+([\d.,]+)"
        amount_str = find(cost_pattern, default="0", source_text=cost_section)
        amount_float = parse_amount(amount_str)

        # Hanya tambahkan baris jika biaya ditemukan
        if amount_float > 0:
            rows.append({
                "file": filename,
                "invoice_number": invoice_number,
                "sender": sender,
                "etd_eta": etd_eta,
                "port_loading": port_loading,
                "port_discharge": port_discharge,
                "invoice_date": invoice_date,
                "stt_number": stt_number,
                "gross_weight_kg": gross_weight_kg,
                "volume_cbm": volume_cbm,
                "cost_type": code,
                "amount": amount_float
            })

    # --- Pembuatan File Excel di Memori ---
    if not rows:
        raise ValueError("Tidak ada data biaya yang dapat diekstrak. Periksa format PDF.")

    df_long = pd.DataFrame(rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_long.to_excel(writer, index=False, sheet_name='InvoiceData')
        
        # Menerapkan styling ke file Excel
        ws = writer.sheets['InvoiceData']
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="4F81BD")
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            # Menentukan lebar kolom berdasarkan konten terpanjang
            for cell in col:
                try:
                    if cell.value:
                        # Menambah sedikit padding
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 4

        # Meratakan angka ke kanan
        for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
            cell = row[-1] # Kolom terakhir (amount)
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = '#,##0.00'
        
        ws.freeze_panes = "A2"  # Membekukan baris header
        ws.auto_filter.ref = ws.dimensions  # Menambah filter otomatis

    return output.getvalue()


def handler(event, context):
    """Fungsi handler yang dipanggil oleh Netlify saat ada request."""
    try:
        # Netlify mengirim data form sebagai string base64 di body.
        # Kode di bawah ini mem-parsingnya secara manual.
        content_type = event['headers'].get('content-type', '')
        if 'multipart/form-data' not in content_type:
            raise ValueError("Tipe konten tidak valid, harus multipart/form-data")

        boundary = content_type.split('boundary=')[1]
        body_decoded = base64.b64decode(event['body'])
        
        # Ekstrak konten file PDF dari body request
        file_part_start = body_decoded.find(b'Content-Type: application/pdf\r\n\r\n')
        if file_part_start == -1:
            raise ValueError("Konten PDF tidak ditemukan dalam request.")
            
        file_content_start = file_part_start + len(b'Content-Type: application/pdf\r\n\r\n')
        file_part_end = body_decoded.find(b'\r\n--' + boundary.encode(), file_content_start)
        pdf_content = body_decoded[file_content_start:file_part_end]

        # Ekstrak nama file asli
        filename_match = re.search(b'filename="([^"]+)"', body_decoded)
        filename = filename_match.group(1).decode() if filename_match else "unknown.pdf"

        # Proses PDF dan dapatkan konten Excel
        excel_data = process_pdf_to_excel(pdf_content, filename)

        # Kembalikan file Excel sebagai respons
        return {
            "statusCode": 200,
            "headers": {
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Content-Disposition": f"attachment; filename=\"parsed_{os.path.splitext(filename)[0]}.xlsx\""
            },
            "body": base64.b64encode(excel_data).decode('utf-8'),
            "isBase64Encoded": True
        }
    except Exception as e:
        # Mengembalikan pesan error dalam format JSON untuk debugging di frontend
        error_message = f"Terjadi kesalahan di server: {str(e)}"
        return {
            "statusCode": 500,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps({"error": error_message})
        }
