import os
import pandas as pd

def read_excel_to_dict_list(file_path):
    """
    处理表格数据放入对应的容器，将表格数据存储为列表，每行的数据存储为字典
    返回值为表格中存储的字典
    """
    data = pd.read_excel(file_path, engine='openpyxl')
    dict_list = []
    for index, row in data.iterrows():
        row_dict = row.to_dict()
        dict_list.append(row_dict)
    return dict_list

    

def create_nested_folders(dict_list):
    """
    生成文件夹序列
    """
    inschoolfolder = '校内'
    outsideschoolfolder = '校外'
    # os.makedirs('哈尔滨职业技术大学\校内', exist_ok=True)
    # os.makedirs('哈尔滨职业技术大学\校外', exist_ok=True)

    for item in dict_list:
        if item in dict_list:
            #以下步骤为了去掉字符串中的空格字符
            colleagename = str(item['开课院系']).strip()
            teachername = str(item['上课教师']).strip()
            teachertype = str(item['是否外聘']).strip()
            coursename = str(item['课程名称']).strip()
            parent_folder = '哈尔滨职业技术大学'
            if teachertype == '是' or teachertype == '校内':
                subfolder_path = os.path.join(parent_folder, str('校内\\'+colleagename+str('\\')+teachername)+str('\\')+coursename)
                os.makedirs(subfolder_path, exist_ok=True)
            elif teachertype == '否' or teachertype == '校外':
                subfolder_path = os.path.join(parent_folder, str('校外\\'+colleagename+str('\\')+teachername)+str('\\')+coursename)
                os.makedirs(subfolder_path, exist_ok=True)

dict_list = read_excel_to_dict_list('23-24-2所有任务.xlsx')
create_nested_folders(dict_list=dict_list)