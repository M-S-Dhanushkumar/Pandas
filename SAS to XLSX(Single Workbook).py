import pandas as pd
import zipfile
import os
import shutil

# Specify the path to the SAS zip file
sas_zip_file_path = '/content/Input_Zip/LIVE_ABBOTTINDIA.ECASELINK.COM_DIEN-122-0195_25JUL2023_002726_DEFAULT.zip'

# Specify the output Excel file name (without extension)
output_excel_file = '/content/Output_xlsx'

# Create a directory to store extracted SAS datasets
os.makedirs('unzipped_data', exist_ok=True)

# Unzip the SAS zip file
with zipfile.ZipFile(sas_zip_file_path, 'r') as zip_ref:
    zip_ref.extractall('unzipped_data')

    # List the extracted files
    extracted_files = zip_ref.namelist()

# Create an Excel writer
excel_writer = pd.ExcelWriter(f'{output_excel_file}.xlsx', engine='xlsxwriter')

# Loop through extracted files and convert to Excel format
for file_name in extracted_files:
    # Check if the file has the .sas7bdat extension
    if file_name.endswith('.sas7bdat'):
        # Read the SAS dataset using sas7bdat
        sas_data = pd.read_sas(f'unzipped_data/{file_name}', format='sas7bdat')

        # Convert bytes to string data by decoding
        sas_data = sas_data.applymap(lambda x: x.decode('utf-8') if isinstance(x, bytes) else x)
        
        # Write the data to the Excel file
        sas_data.to_excel(excel_writer, sheet_name=file_name[:-8], index=False)

        # Save the Excel file
        excel_writer.save()

print("Conversion complete. Your Excel file is ready!")

# Clean up: remove the unzipped directory and all its contents
if os.path.exists('unzipped_data'):
    shutil.rmtree('unzipped_data')
