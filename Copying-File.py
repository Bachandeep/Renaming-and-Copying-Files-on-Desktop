import os
import shutil
import pandas as pd

# === CHANGE THIS to your actual Excel file path ===
excel_file = r"File Path"

desktop = os.path.join(os.path.expanduser("~"), "Desktop")
destination_folder = os.path.join(desktop, "2021(Reg)")

os.makedirs(destination_folder, exist_ok=True)

df = pd.read_excel(excel_file)

for index, row in df.iterrows():
    file_name = row['File_Name']
    file_path = row['File_Path']

    source_file = os.path.join(file_path, file_name)

    if os.path.isfile(source_file):
        try:
            shutil.copy(source_file, destination_folder)
            print(f"Copied: {source_file}")
        except Exception as e:
            print(f"Error copying {source_file}: {e}")
    else:
        print(f"File not found: {source_file}")

print("✅ All possible files copied to Desktop\\2021")
