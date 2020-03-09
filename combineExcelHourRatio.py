# -*- coding:utf-8 -*-
'''
本工具用于合并多个Excel中的多个sheet（具有相同的表头）至一个Excel中的一个sheet
'''
import xlrd, xlsxwriter, os
# from collections import deque

# 获取多个原始Excel列表
def get_source_file_list(file_path):
    all_file_lists = []
    f_list = os.listdir(file_path)
    for file_name in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(file_name)[1] == '.xlsx':
            file_name = file_path + '\\' + file_name
            all_file_lists.append(file_name)
    return all_file_lists

# 初始化一个目标Excel
def init_excel(destination_file_name):
    #   定义一个目标excel
    end_excel = xlsxwriter.Workbook(destination_file_name)
    #   添加一个sheet
    end_excel.add_worksheet('sheet1')
    return end_excel

# 打开一个Excel文件
def open_xls(file):
    try:
        fh = xlrd.open_workbook(file)
        return fh
    except Exception as e:
        print("打开文件错误：" + e)

# 根据excel名以及第几个标签信息就可以得到具体标签的内容
def get_file_value(filename, sheetnum):
    rvalue = []
    fh = open_xls(filename)
    sheet = fh.sheets()[sheetnum]
    row_num = sheet.nrows
    for row in range(0, row_num):
        rvalue.append(sheet.row_values(row))
    return rvalue

# 将列表file_sheet_value的内容写入目标excel
def write_to_end_excel(file_name, end_excel, sheet_value, num_sheet, num_row):
    #打开目标Excel的sheet
    end_xls_sheet = end_excel.get_worksheet_by_name('sheet1')
    num = num_sheet
    num1 = num_row
    for row_info in sheet_value:
        num1 += 1
        num2 = -1
        for colum_info in row_info:
            num2 += 1
            # print(num,num1,num2,sheet3)
            # 在第num1行的第num2列写入sheet3的内容
            end_xls_sheet.write(num1, num2, colum_info)
            end_xls_sheet.write(num1, num2 + 1, file_name)
    return num, num1

# 合并Excel
def combine_excel(file_path, end_xls, log_file):
    # 写记录文件
    log_txt = open(log_file, 'w+')
    # 获取所有在同一个文件夹下的原始文件
    allxls = get_source_file_list(file_path)
    # 初始化一个目标excel
    endxls = init_excel(end_xls)
    num_sheet = -1
    num_row = -1
    #   循环所有的原始文件，保存到一个目标excel
    for file_name in allxls:
        print("正在读取" + file_name)
        print("正在读取" + file_name, file=log_txt)
        # 打开原始文件
        file_fh = open_xls(file_name)
        # 获取原始文件的所有sheet
        file_sheet = file_fh.sheets()
        file_sheet_num = len(file_sheet)
        # 针对每一个sheet
        for sheet_num in range(0, file_sheet_num):
            print("正在读取" + file_name + "的第" + str(sheet_num + 1) + "个标签...", file=log_txt)
            # 获取sheet中数据
            sheet_value = get_file_value(file_name, sheet_num)
            # 写入
            num_sheet, num_row = write_to_end_excel(file_name, endxls, sheet_value, num_sheet, num_row)
    #   关闭
    log_txt.close()
    endxls.close()

def start():
    # 源excel文件夹
    file_path = "F:\\工时占比2019"
    # 目标excel
    end_xls = "F:\\工时占比2019\\结果\\2019年1月至5月工时占比.xlsx"
    log_file = "F:\\工时占比2019\\结果\\日志.txt"
    combine_excel(file_path, end_xls, log_file)

start()