import pandas as pd
import os
import sys
from datetime import datetime

def merge_folder(folder_path):
    if not os.path.isdir(folder_path):
        print("Folder tidak ditemukan.")
        sys.exit(1)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(folder_path, f"merged_{timestamp}.xlsx")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

        for filename in os.listdir(folder_path):
            if filename.endswith((".xlsx", ".xls")) and not filename.startswith("~$"):

                file_path = os.path.join(folder_path, filename)

                try:
                    all_sheets = pd.read_excel(file_path, sheet_name=None)

                    for sheet_name, df in all_sheets.items():

                        new_sheet_name = f"{os.path.splitext(filename)[0]}_{sheet_name}"
                        new_sheet_name = new_sheet_name[:31]

                        counter = 1
                        base_name = new_sheet_name

                        while new_sheet_name in writer.book.sheetnames:
                            new_sheet_name = f"{base_name[:28]}_{counter}"
                            counter += 1

                        df.to_excel(writer, sheet_name=new_sheet_name, index=False)

                except Exception as e:
                    print(f"Gagal baca {filename}: {e}")

    print(f"\n✅ Merge selesai!")
    print(f"File output: {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Cara pakai:")
        print("python merge_excel_cli.py /path/ke/folder")
        sys.exit(1)

    merge_folder(sys.argv[1])

