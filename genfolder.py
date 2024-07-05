import os
import pandas as pd

def read_excel_to_dict_list(file_path):
    """
    处理表格数据放入对应的容器，将表格数据存储为列表，每行的数据存储为字典
    """
    data = pd.read_excel(file_path, engine='openpyxl')
    dict_list = []
    for index, row in data.iterrows():
        row_dict = row.to_dict()
        dict_list.append(row_dict)
    return dict_list

    

def create_nested_folders(dict_list):
    inschoolfolder = '校内'
    outsideschoolfolder = '校外'
    # os.makedirs('哈尔滨职业技术大学\校内', exist_ok=True)
    # os.makedirs('哈尔滨职业技术大学\校外', exist_ok=True)

    for item in dict_list:
        if item in dict_list:
            colleagename = str(item['开课院系']).strip()
            teachername = str(item['上课教师']).strip()
            teachertype = str(item['是否外聘']).strip()
            coursename = str(item['课程名称']).strip()
            parent_folder = '哈尔滨职业技术大学'
            if item['是否外聘'] == '是' or item['是否外聘'] == '校内':
                subfolder_path = os.path.join(parent_folder, str('校内\\'+colleagename+str('\\')+teachername)+str('\\')+coursename)
                os.makedirs(subfolder_path, exist_ok=True)
            elif item['是否外聘'] == '否' or item['是否外聘'] == '校外':
                subfolder_path = os.path.join(parent_folder, str('校外\\'+colleagename+str('\\')+teachername)+str('\\')+coursename)
                os.makedirs(subfolder_path, exist_ok=True)



    

def create_multilevel_folder(folder_path):
    """
    生成文件夹序列
    """
    os.makedirs(folder_path)
    # try:
    #     os.makedirs(folder_path)
    #     print(f"多层文件夹 '{folder_path}' 已创建")
    # except FileExistsError:
    #     print(f"多层文件夹 '{folder_path}' 已存在")
    # except OSError as e:
    #     print(f"创建多层文件夹时发生错误: {e}")

# # 假设 Excel 文件名为 'data.xlsx'
# data = pd.read_excel('data.xlsx')

# for index, row in data.iterrows():
#     create_nested_folders(row)


dict_list = read_excel_to_dict_list('23-24-2所有任务.xlsx')
create_nested_folders(dict_list=dict_list)