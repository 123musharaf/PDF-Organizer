import os
import pandas as pd
import shutil
from flask import Flask, request, render_template, send_file
from pathlib import Path
from zipfile import ZipFile

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'  # Temporary folder on server
DOWNLOAD_FOLDER = 'downloads'  # Temporary folder on server
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def organize_supplier_files(excel_folder_path, pdf_folder_path, output_base_path):
    os.makedirs(output_base_path, exist_ok=True)
    
    excel_files = [f for f in os.listdir(excel_folder_path) 
                   if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    print(f"Found {len(excel_files)} Excel files")
    
    pdf_files = [f for f in os.listdir(pdf_folder_path) if f.endswith('.pdf')]
    print(f"Found {len(pdf_files)} PDF files")
    
    for excel_file in excel_files:
        try:
            excel_path = os.path.join(excel_folder_path, excel_file)
            df = pd.read_excel(excel_path)
            
            if 'Supplier' not in df.columns or 'Part Description' not in df.columns:
                print(f"Skipping {excel_file}: Missing 'Supplier' or 'Part Description' columns")
                continue
            
            supplier = str(df['Supplier'].iloc[0]).strip()
            if not supplier or supplier == 'nan':
                print(f"Skipping {excel_file}: Empty supplier name")
                continue
                
            print(f"\nProcessing {excel_file} for supplier: {supplier}")
            
            supplier_folder = os.path.join(output_base_path, supplier)
            os.makedirs(supplier_folder, exist_ok=True)
            
            dest_excel = os.path.join(supplier_folder, excel_file)
            shutil.copy2(excel_path, dest_excel)
            print(f"Copied {excel_file} to {supplier}")
            
            part_descriptions = df['Part Description'].dropna().astype(str).tolist()
            print(f"Found {len(part_descriptions)} part descriptions")
            
            for part_desc in part_descriptions:
                part_desc = part_desc.strip()
                for pdf_file in pdf_files:
                    pdf_name = os.path.splitext(pdf_file)[0]
                    if part_desc in pdf_name:
                        pdf_source = os.path.join(pdf_folder_path, pdf_file)
                        pdf_dest = os.path.join(supplier_folder, pdf_file)
                        try:
                            shutil.copy2(pdf_source, pdf_dest)
                            print(f"Copied {pdf_file} to {supplier}")
                        except Exception as e:
                            print(f"Error copying {pdf_file}: {str(e)}")
                
        except Exception as e:
            print(f"Error processing {excel_file}: {str(e)}")
    
    zip_path = os.path.join(output_base_path, "organized_suppliers.zip")
    with ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(output_base_path):
            for file in files:
                if file != "organized_suppliers.zip":
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, output_base_path))
    
    return zip_path

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'excel_files' not in request.files or 'pdf_files' not in request.files:
            return "Please upload both Excel and PDF files", 400
        
        excel_files = request.files.getlist('excel_files')
        pdf_files = request.files.getlist('pdf_files')
        
        shutil.rmtree(app.config['UPLOAD_FOLDER'], ignore_errors=True)
        shutil.rmtree(app.config['DOWNLOAD_FOLDER'], ignore_errors=True)
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
        
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'excel')
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], 'pdf')
        os.makedirs(excel_path, exist_ok=True)
        os.makedirs(pdf_path, exist_ok=True)
        
        for file in excel_files:
            if file and file.filename.endswith(('.xlsx', '.xls')):
                file.save(os.path.join(excel_path, file.filename))
        
        for file in pdf_files:
            if file and file.filename.endswith('.pdf'):
                file.save(os.path.join(pdf_path, file.filename))
        
        zip_path = organize_supplier_files(excel_path, pdf_path, app.config['DOWNLOAD_FOLDER'])
        
        return send_file(zip_path, as_attachment=True, download_name="organized_suppliers.zip")
    
    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)