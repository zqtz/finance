import os
import openpyxl
from openpyxl import load_workbook

def excel_file_list(path:str) -> list:
    # lists = os.listdir('F:\\目标与计划\\财务\\learn\\python-openpyxl-files\\video-file\\001 访问目录下所有Excel文件\\4月报表')
    result_list = []
    # 对根路径的每个文件进行迭代
    for each_entry in os.scandir(path):
        # 如果该文件是一个文件夹/目录
        if each_entry.is_dir():
            result_list += excel_file_list(each_entry.path)
        # 如果该文件是一个普通文件,判断一下是不是excel文件,然后将其加入results_list
        else:
            each_entry.path.endswith('.xlsx')
            result_list.append(each_entry.path)
    return result_list

def excel_file_list_iter(path:str) -> list:
    #构造两个列表
    result_list = []
    stack = []
    stack.append(path)
    while len(stack) != 0:
        current_dir = stack.pop()
        if os.path.isdir(current_dir):
            for each_entry in os.scandir(current_dir):
                if each_entry.is_dir():
                    stack.append(each_entry.path)
                else:
                    if each_entry.path.endswith('.xlsx'):
                        result_list.append(each_entry.path)

        else:
            if current_dir.path.endswith('.xlsx'):
                result_list.append(current_dir.path)
    # stack为空
    return result_list




# for i in excel_file_list('F:\\目标与计划\\财务\\learn\\python-openpyxl-files\\video-file\\001 访问目录下所有Excel文件\\4月报表'):
#     print(i)
# print('*'*1000)

for each_path in excel_file_list_iter('F:\\目标与计划\\财务\\learn\\python-openpyxl-files\\video-file\\001 访问目录下所有Excel文件\\4月报表'):
    try:
        wb = load_workbook(each_path)
        print(wb.active['A3'].value)
        wb.close()
    except:
        continue