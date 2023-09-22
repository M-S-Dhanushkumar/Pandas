import pandas as pd
import zipfile
import os
import shutil

# Specify the path to the SAS zip file
sas_zip_file_path = '/content/Path_Zip File.zip'

# Create a directory to store extracted SAS datasets
os.makedirs('unzipped_data', exist_ok=True)

# Unzip the SAS zip file
with zipfile.ZipFile(sas_zip_file_path, 'r') as zip_ref:
    zip_ref.extractall('unzipped_data')

# List the extracted files
extracted_files = os.listdir('unzipped_data')

# Create a directory to store individual Excel files
os.makedirs('xlsx_data', exist_ok=True)

# Define a function to decode bytes to strings
def decode_bytes(x):
    if isinstance(x, bytes):
        try:
            return x.decode('utf-8')
        except UnicodeDecodeError:
            return str(x)  # If decoding fails, convert to a string representation
    elif isinstance(x, float):
        return str(x)  # Convert floats to strings
    else:
        return x  # Keep other types of values as is

# Loop through extracted files and convert to Excel format
for file_name in extracted_files:
    # Check if the file has the .sas7bdat extension
    if file_name.endswith('.sas7bdat'):
        # Read the SAS dataset using sas7bdat
        sas_data = pd.read_sas(f'unzipped_data/{file_name}', format='sas7bdat')
        
        # Determine the output Excel file name without modifications
        excel_file_name = f'xlsx_data/{file_name[:-8]}.xlsx'

        # Apply the decoding function to your DataFrame
        sas_data = sas_data.applymap(decode_bytes)

        # Create an Excel writer for the individual file
        excel_writer = pd.ExcelWriter(excel_file_name, engine='xlsxwriter')
        
        # Write the data to the Excel file with a specified sheet name (e.g., 'Sheet1')
        sas_data.to_excel(excel_writer, sheet_name='Sheet1', index=False)
        
        # Save the Excel file
        excel_writer.save()

print("Conversion complete. SAS datasets have been exported to individual Excel (.xlsx) files.")

# Clean up: remove the unzipped directory and all its contents
if os.path.exists('unzipped_data'):
    shutil.rmtree('unzipped_data')
