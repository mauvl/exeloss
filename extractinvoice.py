def process_invoice(zip_path):
    import zipfile
    import fitz
    import re
    import os
    import csv
    import shutil

    output_csv = "hasil_invoice.csv"
    extract_folder = "pdf_temp"
    os.makedirs(extract_folder, exist_ok=True)

    # Bersihkan folder
    for file in os.listdir(extract_folder):
        file_path = os.path.join(extract_folder, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_folder)

    pattern_invoice = re.compile(r"(?:Invoice Number|Nomor Invoice).*?(\d{8,12})", re.DOTALL)
    pattern_customer_id = re.compile(r"ID Pelanggan\s*/\s*Customer ID\s*\n?\s*(\d+)")
    pattern_customer_name = re.compile(r"Nama\s+Account\s*/\s+Account Name\s*\n?\s*(.+?)\n")

    hasil = []

    for filename in os.listdir(extract_folder):
        if filename.lower().endswith(".pdf"):
            path_pdf = os.path.join(extract_folder, filename)
            doc = fitz.open(path_pdf)
            text = ""
            for page in doc:
                text += page.get_text()

            invoice_match = pattern_invoice.search(text)
            customer_id_match = pattern_customer_id.search(text)
            customer_name_match = pattern_customer_name.search(text)

            invoice_number = invoice_match.group(1) if invoice_match else "N/A"
            customer_id = customer_id_match.group(1) if customer_id_match else "N/A"
            customer_name = customer_name_match.group(1) if customer_name_match else "N/A"

            hasil.append({
                "Filename": filename,
                "Customer ID": customer_id,
                "Invoice Number": invoice_number,
                "Customer Name": customer_name,
                "PDF Filename": filename
            })

    # Simpan ke CSV
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file_csv:
        writer = csv.writer(file_csv)
        writer.writerow(["Filename", "Customer ID", "Invoice Number", "Customer Name", "PDF Filename"])
        for row in hasil:
            writer.writerow([row["Filename"], row["Customer ID"], row["Invoice Number"], row["Customer Name"], row["PDF Filename"]])

    return output_csv, hasil
