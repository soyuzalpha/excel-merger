import os
from openpyxl import load_workbook, Workbook

def merge_folder():
    folder_path = os.getcwd()
    new_wb = Workbook()
    new_ws = new_wb.active
    new_ws.title = "Merged"

    current_row = 1
    file_count = 0

    for filename in os.listdir(folder_path):
        if (
            filename.endswith(".xlsx")
            and not filename.startswith("~$")
            and not filename.startswith("merged_")
        ):
            try:
                wb = load_workbook(os.path.join(folder_path, filename), data_only=True)
                ws = wb.active

                for row in ws.iter_rows(values_only=True):
                    for col_idx, value in enumerate(row, start=1):
                        new_ws.cell(row=current_row, column=col_idx, value=value)
                    current_row += 1

                file_count += 1

            except Exception as e:
                print(f"Gagal baca {filename}: {e}")

    if file_count == 0:
        print("Tidak ada file yang berhasil di-merge.")
        return

    output_file = os.path.join(folder_path, "merged_result.xlsx")
    new_wb.save(output_file)

    print(f"Berhasil merge {file_count} file.")
    print(f"Hasil: {output_file}")

