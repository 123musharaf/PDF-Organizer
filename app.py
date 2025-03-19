# import os
# import pandas as pd
# import shutil
# from flask import Flask, request, render_template, send_file
# from pathlib import Path
# from zipfile import ZipFile

# app = Flask(__name__)

# # Configuration
# DOWNLOAD_FOLDER = r'C:\Users\LENOVO\Desktop\supplier_organizer_web\downloads'
# app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# # Ensure download folder exists
# os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# def organize_supplier_files(folder_path, output_base_path):
#     # Create output base directory if it doesn't exist
#     os.makedirs(output_base_path, exist_ok=True)

#     # Get all Excel files (excluding temporary files)
#     excel_files = [f for f in os.listdir(folder_path) 
#                    if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
#     print(f"Found {len(excel_files)} Excel files")

#     # Use the same folder for PDFs
#     pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
#     print(f"Found {len(pdf_files)} PDF files")

#     for excel_file in excel_files:
#         try:
#             # Read Excel file
#             excel_path = os.path.join(folder_path, excel_file)
#             df = pd.read_excel(excel_path)

#             # Verify columns
#             if 'Supplier' not in df.columns or 'Part Description' not in df.columns:
#                 print(f"Skipping {excel_file}: Missing 'Supplier' or 'Part Description' columns")
#                 continue

#             # Get supplier name (same for all rows)
#             supplier = str(df['Supplier'].iloc[0]).strip()
#             if not supplier or supplier == 'nan':
#                 print(f"Skipping {excel_file}: Empty supplier name")
#                 continue

#             print(f"\nProcessing {excel_file} for supplier: {supplier}")

#             # Create supplier folder
#             supplier_folder = os.path.join(output_base_path, supplier)
#             os.makedirs(supplier_folder, exist_ok=True)

#             # Copy Excel file to supplier folder
#             dest_excel = os.path.join(supplier_folder, excel_file)
#             shutil.copy2(excel_path, dest_excel)
#             print(f"Copied {excel_file} to {supplier}")

#             # Process all part descriptions
#             part_descriptions = df['Part Description'].dropna().astype(str).tolist()
#             print(f"Found {len(part_descriptions)} part descriptions")

#             # Find and copy matching PDFs from the same folder
#             for part_desc in part_descriptions:
#                 part_desc = part_desc.strip()
#                 for pdf_file in pdf_files:
#                     pdf_name = os.path.splitext(pdf_file)[0]
#                     if part_desc in pdf_name:
#                         pdf_source = os.path.join(folder_path, pdf_file)
#                         pdf_dest = os.path.join(supplier_folder, pdf_file)
#                         try:
#                             shutil.copy2(pdf_source, pdf_dest)
#                             print(f"Copied {pdf_file} to {supplier}")
#                         except Exception as e:
#                             print(f"Error copying {pdf_file}: {str(e)}")

#         except Exception as e:
#             print(f"Error processing {excel_file}: {str(e)}")

#     # Create ZIP file of output
#     zip_path = os.path.join(output_base_path, "organized_suppliers.zip")
#     with ZipFile(zip_path, 'w') as zipf:
#         for root, _, files in os.walk(output_base_path):
#             for file in files:
#                 if file != "organized_suppliers.zip":
#                     file_path = os.path.join(root, file)
#                     zipf.write(file_path, os.path.relpath(file_path, output_base_path))

#     return zip_path

# @app.route('/', methods=['GET', 'POST'])
# def index():
#     if request.method == 'POST':
#         # Get folder path from form
#         folder_path = request.form.get('folder_path')

#         if not folder_path or not os.path.exists(folder_path):
#             return "Please provide a valid folder path", 400

#         # Clear previous downloads
#         shutil.rmtree(app.config['DOWNLOAD_FOLDER'], ignore_errors=True)
#         os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

#         # Process files from the specified folder
#         zip_path = organize_supplier_files(folder_path, app.config['DOWNLOAD_FOLDER'])

#         return send_file(zip_path, as_attachment=True, download_name="organized_suppliers.zip")

#     return render_template('index.html')

# if __name__ == '_main_':
#     app.run(debug=True)













import os
import pandas as pd
import shutil
from flask import Flask, request, render_template, send_file
from pathlib import Path
from zipfile import ZipFile

app = Flask(__name__)

# Configuration
DOWNLOAD_FOLDER = 'downloads'
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# Ensure download folder exists
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def organize_supplier_files(folder_path, output_base_path):
    # Create output base directory if it doesn't exist
    os.makedirs(output_base_path, exist_ok=True)
    
    # Get all Excel files (excluding temporary files)
    excel_files = [f for f in os.listdir(folder_path) 
                   if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    print(f"Found {len(excel_files)} Excel files")
    
    # Use the same folder for PDFs
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    print(f"Found {len(pdf_files)} PDF files")
    
    for excel_file in excel_files:
        try:
            # Read Excel file
            excel_path = os.path.join(folder_path, excel_file)
            df = pd.read_excel(excel_path)
            
            # Verify columns
            if 'Supplier' not in df.columns or 'Part Description' not in df.columns:
                print(f"Skipping {excel_file}: Missing 'Supplier' or 'Part Description' columns")
                continue
            
            # Get supplier name (same for all rows)
            supplier = str(df['Supplier'].iloc[0]).strip()
            if not supplier or supplier == 'nan':
                print(f"Skipping {excel_file}: Empty supplier name")
                continue
                
            print(f"\nProcessing {excel_file} for supplier: {supplier}")
            
            # Create supplier folder
            supplier_folder = os.path.join(output_base_path, supplier)
            os.makedirs(supplier_folder, exist_ok=True)
            
            # Copy Excel file to supplier folder
            dest_excel = os.path.join(supplier_folder, excel_file)
            shutil.copy2(excel_path, dest_excel)
            print(f"Copied {excel_file} to {supplier}")
            
            # Process all part descriptions
            part_descriptions = df['Part Description'].dropna().astype(str).tolist()
            print(f"Found {len(part_descriptions)} part descriptions")
            
            # Find and copy matching PDFs from the same folder
            for part_desc in part_descriptions:
                part_desc = part_desc.strip()
                for pdf_file in pdf_files:
                    pdf_name = os.path.splitext(pdf_file)[0]
                    if part_desc in pdf_name:
                        pdf_source = os.path.join(folder_path, pdf_file)
                        pdf_dest = os.path.join(supplier_folder, pdf_file)
                        try:
                            shutil.copy2(pdf_source, pdf_dest)
                            print(f"Copied {pdf_file} to {supplier}")
                        except Exception as e:
                            print(f"Error copying {pdf_file}: {str(e)}")
                
        except Exception as e:
            print(f"Error processing {excel_file}: {str(e)}")
    
    # Create ZIP file of output
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
        # Get folder path from form
        folder_path = request.form.get('folder_path')
        
        if not folder_path or not os.path.exists(folder_path):
            return "Please provide a valid folder path", 400
        
        # Clear previous downloads
        shutil.rmtree(app.config['DOWNLOAD_FOLDER'], ignore_errors=True)
        os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
        
        # Process files from the specified folder
        zip_path = organize_supplier_files(folder_path, app.config['DOWNLOAD_FOLDER'])
        
        return send_file(zip_path, as_attachment=True, download_name="organized_suppliers.zip")
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)