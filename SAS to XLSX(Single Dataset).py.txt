import pandas as pd

# Specify the path to the SAS dataset file (.sas7bdat)
sas_dataset_path = '/path/to/your/dataset.sas7bdat'

# Specify the output Excel file name (with .xlsx extension)
output_excel_file = '/path/to/your/output_dataset.xlsx'

# Read the SAS dataset using sas7bdat
sas_data = pd.read_sas(sas_dataset_path, format='sas7bdat')

# Write the data to the Excel file
sas_data.to_excel(output_excel_file, index=False)

print(f"Conversion complete. SAS dataset has been exported to '{output_excel_file}'.")
