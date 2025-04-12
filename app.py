import os
import pandas as pd
import shutil
from flask import Flask, request, render_template, send_file
from zipfile import ZipFile

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def organize_supplier_files(excel_folder_path, pdf_folder_path, output_base_path):
    os.makedirs(output_base_path, exist_ok=True)
    
    excel_files = [f for f in os.listdir(excel_folder_path) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    pdf_files = [f for f in os.listdir(pdf_folder_path) if f.endswith('.pdf')]
    pdf_file_names = [os.path.splitext(f)[0].lower() for f in pdf_files]
    unmatched_pdfs = set(pdf_file_names)

    for excel_file in excel_files:
        try:
            df = pd.read_excel(os.path.join(excel_folder_path, excel_file))

            if 'Supplier' not in df.columns or 'Part Description' not in df.columns:
                continue

            grouped = df.groupby('Supplier')

            for supplier, group_df in grouped:
                if not isinstance(supplier, str) or supplier.strip().lower() == 'nan':
                    continue

                supplier = supplier.strip()
                supplier_folder = os.path.join(output_base_path, supplier)
                os.makedirs(supplier_folder, exist_ok=True)

                dest_excel = os.path.join(supplier_folder, f"{supplier}_{excel_file}")
                group_df.to_excel(dest_excel, index=False)

                seen_parts = set()

                for part_desc in group_df['Part Description'].dropna().astype(str).tolist():
                    part_desc_clean = part_desc.strip().lower()
                    found = False

                    for i, pdf_file in enumerate(pdf_files):
                        pdf_name_clean = pdf_file_names[i]

                        if part_desc_clean in pdf_name_clean:
                            pdf_source = os.path.join(pdf_folder_path, pdf_file)

                            filename = f"{part_desc.strip()}.pdf"
                            if part_desc_clean in seen_parts:
                                filename = f"{part_desc.strip()}_SPARE.pdf"
                            else:
                                seen_parts.add(part_desc_clean)

                            pdf_dest = os.path.join(supplier_folder, filename)

                            try:
                                shutil.copy2(pdf_source, pdf_dest)
                                print(f"Copied {pdf_file} as {filename} to {supplier}")
                                found = True
                                unmatched_pdfs.discard(pdf_name_clean)
                            except Exception as e:
                                print(f"Error copying {pdf_file}: {str(e)}")

                    if not found:
                        print(f"Warning: No matching PDF found for part description '{part_desc}'")
        except Exception as e:
            print(f"Error processing {excel_file}: {str(e)}")

    zip_path = os.path.join(output_base_path, "organized_suppliers.zip")
    with ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(output_base_path):
            for file in files:
                if file != "organized_suppliers.zip":
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, output_base_path))

    return zip_path, list(unmatched_pdfs)

@app.route('/', methods=['GET', 'POST'])
def index():
    unmatched_pdfs = []
    
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

        zip_path, unmatched_pdfs = organize_supplier_files(excel_path, pdf_path, app.config['DOWNLOAD_FOLDER'])

        return render_template('index.html', download_link='/download', unmatched_pdfs=unmatched_pdfs)

    return render_template('index.html', unmatched_pdfs=unmatched_pdfs)

@app.route('/download')
def download():
    zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], "organized_suppliers.zip")
    return send_file(zip_path, as_attachment=True, download_name="organized_suppliers.zip")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
