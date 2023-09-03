import openpyxl
import test
import shutil
import datetime
import time
import os

def generate_new_filename(base_path):
    current_datetime = datetime.datetime.now().strftime('%m月%d日%H时%M分')
    new_filename = f"{current_datetime}.xlsx"
    return os.path.join(base_path, "docs", new_filename)

def extract_links_from_excel(file_name):
    workbook = openpyxl.load_workbook(file_name)
    all_links = []

    for sheet in workbook.worksheets:
        link_column_index = None

        for cell in sheet[2]:
            if cell.value == "链接":
                link_column_index = cell.column
                break

        if link_column_index:
            for row in sheet.iter_rows(min_col=link_column_index, max_col=link_column_index, min_row=3):
                link_cell = row[0]
                if link_cell.value:
                    all_links.append(link_cell.value)

    return all_links


def run_program(callback=None):
    start_time = time.time()

    base_path = "./"
    new_filename = generate_new_filename(base_path)

    source_file = base_path + "all_1.xlsx"

    # 确保源文件存在
    if not os.path.exists(source_file):
        if callback:
            callback(f"Source file '{source_file}' does not exist!")
        else:
            print(f"Source file '{source_file}' does not exist!")
        return

    shutil.copy2(source_file, new_filename)
    if callback:
        callback(f"File copied to: {new_filename}")
    else:
        print(f"File copied to: {new_filename}")

    file_name = new_filename
    links = extract_links_from_excel(file_name)
    test.process_links(file_name, links)

    end_time = time.time()
    elapsed_time = end_time - start_time
    if callback:
        callback(f"Program ran in {elapsed_time:.2f} seconds.")
    else:
        print(f"Program ran in {elapsed_time:.2f} seconds.")

    return new_filename


if __name__ == "__main__":
    new_filename = run_program()
