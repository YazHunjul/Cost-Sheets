import zipfile

excel_path = "resources/Halton Cost Sheet Jan 2025.xlsx"

with zipfile.ZipFile(excel_path, 'r') as zip_ref:
    # List all files in the Excel file
    for file in zip_ref.namelist():
        print(file) 