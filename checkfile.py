import os
import xlsxwriter

def get_folder_info(folder_path):
    workbook = xlsxwriter.Workbook('folder_info.xlsx')
    worksheet = workbook.add_worksheet()

    row = 0
    col = 0

    def process_folder(path, level):
        nonlocal row
        nonlocal col
        folders = [f for f in os.listdir(path) if os.path.isdir(os.path.join(path, f))]
        for folder in folders:
            full_path = os.path.join(path, folder)
            worksheet.write(row, col + level - 1, folder)
            if len([f for f in os.listdir(full_path) if os.path.isdir(os.path.join(full_path, f))]) == 0:
                file_count = len([f for f in os.listdir(full_path) if os.path.isfile(os.path.join(full_path, f))])
                folder_size = sum(os.path.getsize(os.path.join(full_path, f)) for f in os.listdir(full_path) if os.path.isfile(os.path.join(full_path, f)))
                worksheet.write(row, col + level + 1, file_count)
                worksheet.write(row, col + level + 2, folder_size)
            row += 1
            process_folder(full_path, level + 1)

    process_folder(folder_path, 1)

    workbook.close()

folder_path = "哈尔滨职业技术大学"  # 替换为您要检查的文件夹路径
get_folder_info(folder_path)