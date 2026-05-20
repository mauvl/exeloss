import zipfile
import fitz  # PyMuPDF
import re
import os
from openpyxl import Workbook
from openpyxl.styles import numbers

def process_faktur(zip_path):
    output_xlsx = "hasil_faktur_invoice.xlsx"
    extract_folder = "pdf_temp"
    os.makedirs(extract_folder, exist_ok=True)

    # BERSIHIN ISI LAMA DULU BIAR GA KECAMPUR
    for file in os.listdir(extract_folder):
        file_path = os.path.join(extract_folder, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_folder)

    pattern_faktur = re.compile(r"Kode dan Nomor Seri Faktur Pajak\s*:\s*([\d.]+)")
    pattern_invoice = re.compile(r"\(Referensi[:：]?\s*(\d+)\)")
    pattern_customer_fallback = re.compile(r"Pembeli.*?Nama\s*:\s*(.*?)\n", re.DOTALL | re.IGNORECASE)

    # Init workbook Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Faktur Data"
    ws.append(["Customer Name", "Invoice Number", "Faktur Number"])

    hasil_data = []
    gagal_parse = []

    for filename in os.listdir(extract_folder):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(extract_folder, filename)
            doc = fitz.open(file_path)
            full_text = "\n".join([page.get_text() for page in doc])

            faktur_match = pattern_faktur.search(full_text)
            invoice_match = pattern_invoice.search(full_text)
            customer_name_match = pattern_customer_fallback.search(full_text)

            faktur_number = faktur_match.group(1).strip() if faktur_match else "N/A"
            invoice_number = invoice_match.group(1).strip() if invoice_match else "N/A"
            customer_name = customer_name_match.group(1).strip() if customer_name_match else "N/A"

            ws.append([customer_name, invoice_number, faktur_number])
            hasil_data.append({
                "Customer Name": customer_name,
                "Invoice Number": invoice_number,
                "Faktur Number": faktur_number
            })

            if "N/A" in [customer_name, invoice_number, faktur_number]:
                gagal_parse.append(filename)

    # Set semua kolom jadi format teks
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=3):
        for cell in row:
            cell.number_format = numbers.FORMAT_TEXT

    wb.save(output_xlsx)

    return output_xlsx, hasil_data
