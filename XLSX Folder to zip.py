import zipfile
import os

# Directory containing the Excel (.xlsx) files
xlsx_data_directory = 'xlsx_data'

# Name of the output zip file
output_zip_file = 'output_excel_files.zip'

# Create a zip file containing the Excel files
with zipfile.ZipFile(output_zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for root, _, files in os.walk(xlsx_data_directory):
        for file in files:
            file_path = os.path.join(root, file)
            zipf.write(file_path, os.path.relpath(file_path, xlsx_data_directory))

print(f"Excel files in '{xlsx_data_directory}' have been compressed into '{output_zip_file}'.")
