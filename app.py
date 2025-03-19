import os
import shutil
import pandas as pd
from flask import Flask, request, render_template, send_file
from zipfile import ZipFile

app = Flask(__name__)

# Output folder for organized files
DOWNLOAD_FOLDER = 'downloads'
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get uploaded files from the form
        files = request.files.getlist('folderInput')

        if not files:
            return "Please select a valid folder", 400

        # Create a temporary folder to store uploaded files
        temp_folder = os.path.join(DOWNLOAD_FOLDER, 'uploaded_files')
        os.makedirs(temp_folder, exist_ok=True)

        # Save files to the temp folder
        for file in files:
            save_path = os.path.join(temp_folder, file.filename)
            file.save(save_path)

        # Process files
        zip_path = organize_supplier_files(temp_folder, DOWNLOAD_FOLDER)

        return send_file(zip_path, as_attachment=True, download_name="organized_suppliers.zip")

    return render_template('index.html')

def organize_supplier_files(folder_path, output_base_path):
    """Processes Excel and PDF files, organizes by supplier, and creates a ZIP."""
    os.makedirs(output_base_path, exist_ok=True)
    
    # Get Excel and PDF files
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

    for excel_file in excel_files:
        try:
            excel_path = os.path.join(folder_path, excel_file)
            df = pd.read_excel(excel_path)

            if 'Supplier' not in df.columns or 'Part Description' not in df.columns:
                continue  # Skip if required columns are missing
            
            supplier = str(df['Supplier'].iloc[0]).strip()
            if not supplier:
                continue

            supplier_folder = os.path.join(output_base_path, supplier)
            os.makedirs(supplier_folder, exist_ok=True)

            shutil.copy2(excel_path, supplier_folder)

            part_descriptions = df['Part Description'].dropna().astype(str).tolist()
            
            for part_desc in part_descriptions:
                part_desc = part_desc.strip()
                for pdf_file in pdf_files:
                    if part_desc in pdf_file:
                        shutil.copy2(os.path.join(folder_path, pdf_file), supplier_folder)

        except Exception as e:
            print(f"Error processing {excel_file}: {str(e)}")

    # Create ZIP of organized files
    zip_path = os.path.join(output_base_path, "organized_suppliers.zip")
    with ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(output_base_path):
            for file in files:
                if file != "organized_suppliers.zip":
                    zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), output_base_path))
    
    return zip_path

if __name__ == '__main__':
    app.run(debug=True)
