import pandas as pd
import zipfile
import os
import shutil

# Specify the path to the SAS zip file
sas_zip_file_path = '/content/(Replace your path).zip'

# Create a directory to store extracted SAS datasets
os.makedirs('unzipped_data', exist_ok=True)

# Unzip the SAS zip file
with zipfile.ZipFile(sas_zip_file_path, 'r') as zip_ref:
    zip_ref.extractall('unzipped_data')

    # List the extracted files
    extracted_files = zip_ref.namelist()

# Create a directory to store the individual Excel files
os.makedirs('excel_files', exist_ok=True)

# Loop through extracted files and convert to separate Excel files
for file_name in extracted_files:
    # Check if the file has the .sas7bdat extension
    if file_name.endswith('.sas7bdat'):
        # Read the SAS dataset using sas7bdat
        sas_data = pd.read_sas(f'unzipped_data/{file_name}', format='sas7bdat')

        # Check if the dataset is not empty
        if not sas_data.empty:
            # Modify the sheet name to remove trailing spaces and periods
            sheet_name = file_name[:-8].strip(" .")

            # Convert bytes to string data by decoding
            sas_data = sas_data.applymap(lambda x: x.decode('utf-8', errors='replace') if isinstance(x, bytes) else x)
        
            # Define the individual Excel file name
            output_excel_file = f'excel_files/{sheet_name}.xlsx'

            # Create an Excel writer for the individual file
            excel_writer = pd.ExcelWriter(output_excel_file, engine='xlsxwriter')

            # Write the data to the individual Excel file
            sas_data.to_excel(excel_writer, sheet_name=sheet_name, index=False)

            # Save the individual Excel file
            excel_writer.save()

# Zip the individual Excel files into a single zip file
shutil.make_archive('/content/Output', 'zip', 'excel_files')

print("Conversion complete. Your individual Excel files have been saved and zipped!")

# Clean up: remove the unzipped directory and all its contents, and the individual Excel files
if os.path.exists('unzipped_data'):
    shutil.rmtree('unzipped_data')

if os.path.exists('excel_files'):
    shutil.rmtree('excel_files')
