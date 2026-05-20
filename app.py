from flask import Flask, render_template, request, send_file, send_from_directory, flash, redirect, url_for, abort, jsonify
import os
from werkzeug.utils import secure_filename
from extractfaktur import process_faktur
from extractinvoice import process_invoice
from datetime import datetime
import datetime
import shutil
import zipfile
import pandas as pd
import json
from urllib.parse import quote

app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ARCHIVE_FOLDER'] = 'archive'
app.config['RESULT_FOLDER'] = 'results'
app.config['PDF_FOLDER'] = 'pdf_temp'

for folder in [app.config['UPLOAD_FOLDER'], app.config['ARCHIVE_FOLDER'], app.config['RESULT_FOLDER'], app.config['PDF_FOLDER']]:
    os.makedirs(folder, exist_ok=True)

ALLOWED_EXTENSIONS = {'zip'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
@app.template_filter('insert_date_separators')
def insert_date_separators(date_str):
    if len(date_str) >= 8:
        return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    return date_str

@app.template_filter('get_file_size')
def get_file_size(filename):
    folder = app.config['RESULT_FOLDER']
    path = os.path.join(folder, filename)

    try:
        size = os.path.getsize(path)
        return f"{size / 1024:.2f} KB"
    except:
        return "Unknown"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        return "No file part", 400

    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        return "No selected file", 400

    file_type = request.form.get('file_type')
    combined_data = []
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    result_basename = f"{file_type}_result_{timestamp}"
    result_filename = f"{result_basename}.xlsx" if file_type == 'faktur' else f"{result_basename}.csv"
    result_filepath = os.path.join(app.config['RESULT_FOLDER'], result_filename)

    shutil.rmtree(app.config['PDF_FOLDER'], ignore_errors=True)
    os.makedirs(app.config['PDF_FOLDER'], exist_ok=True)

    for file in files:
        if not allowed_file(file.filename):
            return f"❌ Invalid file type: {file.filename}. Please upload ZIP files only.", 400

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(file_path)

        if file_type == 'faktur':
            result_file, data = process_faktur(file_path)
        elif file_type == 'invoice':
            result_file, data = process_invoice(file_path)
        else:
            return "Unknown file type", 400

        for item in data:
            if 'PDF Filename' not in item or not item['PDF Filename']:
                item['PDF Filename'] = "N/A"

        combined_data.extend(data)
        shutil.move(file_path, os.path.join(app.config['ARCHIVE_FOLDER'], file.filename))
        shutil.move(result_file, result_filepath)

    create_alternative_formats(result_filepath, result_basename)
    return render_template('result.html', data=combined_data, result_file=result_basename)

def create_alternative_formats(source_path, basename):
    try:
        if source_path.endswith('.csv'):
            df = pd.read_csv(source_path)
        else:
            df = pd.read_excel(source_path)

        if not source_path.endswith('.csv'):
            df.to_csv(os.path.join(app.config['RESULT_FOLDER'], f"{basename}.csv"), index=False)

        if not source_path.endswith('.xlsx'):
            df.to_excel(os.path.join(app.config['RESULT_FOLDER'], f"{basename}.xlsx"), index=False)

        df.to_json(os.path.join(app.config['RESULT_FOLDER'], f"{basename}.json"), orient='records', indent=2)
        df.to_string(os.path.join(app.config['RESULT_FOLDER'], f"{basename}.txt"), index=False)
    except Exception as e:
        print(f"Error creating alternative formats: {e}")

@app.route('/results/<filename>')
def download_result(filename):
    if filename.endswith(('.csv', '.xlsx', '.json', '.txt')):
        return send_from_directory(app.config['RESULT_FOLDER'], filename, as_attachment=True)
    else:
        return send_from_directory(app.config['RESULT_FOLDER'], f"{filename}.csv", as_attachment=True)

@app.route('/pdf/<filename>')
def preview_pdf(filename):
    try:
        return send_from_directory(app.config['PDF_FOLDER'], filename)
    except FileNotFoundError:
        try:
            return send_from_directory(app.config['ARCHIVE_FOLDER'], filename)
        except FileNotFoundError:
            abort(404, description="PDF file not found")

@app.route('/download-pdf/<filename>')
def download_pdf(filename):
    try:
        response = send_from_directory(app.config['PDF_FOLDER'], filename, as_attachment=True)
        response.headers["Content-Disposition"] = f"attachment; filename={quote(filename)}"
        return response
    except FileNotFoundError:
        try:
            response = send_from_directory(app.config['ARCHIVE_FOLDER'], filename, as_attachment=True)
            response.headers["Content-Disposition"] = f"attachment; filename={quote(filename)}"
            return response
        except FileNotFoundError:
            abort(404, description="PDF file not found")

@app.route('/check-pdf/<filename>')
def check_pdf(filename):
    pdf_path = os.path.join(app.config['PDF_FOLDER'], filename)
    archive_path = os.path.join(app.config['ARCHIVE_FOLDER'], filename)
    if os.path.exists(pdf_path) or os.path.exists(archive_path):
        return jsonify({"exists": True})
    return jsonify({"exists": False})

@app.route('/delete-file/<filename>', methods=['DELETE'])
def delete_file(filename):
    try:
        filepath = os.path.join(app.config['RESULT_FOLDER'], filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            return jsonify({"success": True})
        else:
            return jsonify({"success": False, "message": "File not found"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})
    
@app.route('/delete-all-files', methods=['DELETE'])
def delete_all_files():
    try:
        result_folder = app.config['RESULT_FOLDER']
        files = [f for f in os.listdir(result_folder) if f.endswith(('.csv', '.xlsx', '.json', '.txt'))]

        for file in files:
            os.remove(os.path.join(result_folder, file))

        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


@app.route('/archive-month', methods=['POST'])
def archive_month():
    now = datetime.datetime.now()
    prefix = now.strftime('%Y%m')
    files = [f for f in os.listdir(app.config['RESULT_FOLDER']) if f.startswith(prefix)]

    if not files:
        return jsonify({"success": False, "message": "No files to archive this month."})

    archive_name = f"monthly_archive_{prefix}.zip"
    archive_path = os.path.join(app.config['ARCHIVE_FOLDER'], archive_name)

    with zipfile.ZipFile(archive_path, 'w') as archive:
        for file in files:
            file_path = os.path.join(app.config['RESULT_FOLDER'], file)
            archive.write(file_path, arcname=file)

    return jsonify({"success": True, "message": f"Archive {archive_name} created."})

@app.route('/history')
def history():
    files = os.listdir(app.config['RESULT_FOLDER'])
    files = sorted(
        [f for f in files if f.endswith(('.csv', '.xlsx', '.json', '.txt'))],
        key=lambda x: os.path.getmtime(os.path.join(app.config['RESULT_FOLDER'], x)),
        reverse=True
    )
    return render_template('history.html', files=files)



@app.route('/dashboard')
def dashboard():
    # Dictionary untuk menghitung jumlah invoice per perusahaan
    company_count = {}
    original_names = {}

    result_files = [f for f in os.listdir(app.config['RESULT_FOLDER']) 
                    if f.endswith(('.csv', '.xlsx'))]

    for result_file in result_files:
        file_path = os.path.join(app.config['RESULT_FOLDER'], result_file)
        try:
            if result_file.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            # GANTI INI SESUAI NAMA PASTI KOLOM PERUSAHAAN DI FILE 
            possible_columns = ['Nama Perusahaan', 'Customer Name', 'customer', 'Customer']

            company_col = next((col for col in df.columns if col in possible_columns), None)

            if company_col:
                for raw_name in df[company_col]:
                    if pd.notna(raw_name):
                        norm_name = str(raw_name).strip().lower()

                        company_count[norm_name] = company_count.get(norm_name, 0) + 1

                        if norm_name not in original_names:
                            original_names[norm_name] = str(raw_name).strip()

        except Exception as e:
            print(f"Error processing {result_file}: {e}")
            continue


    sorted_companies = sorted(company_count.items(), key=lambda x: x[1], reverse=True)

    labels = [original_names[norm_name] for norm_name, _ in sorted_companies]
    counts = [count for _, count in sorted_companies]

    limit = min(5, len(labels))
    return render_template('dashboard.html', labels=labels, counts=counts, limit=limit)




if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
