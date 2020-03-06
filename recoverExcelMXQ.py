# -*- coding:utf-8 -*-
'''
本工具用于合并多个Excel中的多个sheet（具有相同的表头）至一个Excel中的一个sheet
'''
import xlrd, xlsxwriter, os
# from collections import deque

# 获取多个原始Excel列表
def get_origin_file_list(file_path):
    all_file_lists = []
    f_list = os.listdir(file_path)
    for file_name in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(file_name)[1] == '.xlsx':
            file_name = file_path + '\\' + file_name
            all_file_lists.append(file_name)
    return all_file_lists

# 打开一个Excel文件
def open_xls(file):
    try:
        fh = xlrd.open_workbook(file)
        return fh
    except Exception as e:
        print("打开文件错误：" + e)

# 将列表file_sheet_value的内容写入目标excel
def get_work_part_value(origin_xls, sheet_num):
    fh = open_xls(origin_xls)
    sheet = fh.sheets()[sheet_num]
    work_part_value = sheet.row_values(5)[3]
    return work_part_value

def write_work_part_value(target_xls,target_xls_sheet_name, work_part_value):
    # 打开目标Excel的sheet
    end_xls_sheet = target_xls.get_worksheet_by_name(target_xls_sheet_name)
    end_xls_sheet.write(4, 30, work_part_value)

def write_to_target_excel(origin_xls,sheet_num,target_xls,target_xls_sheet_name):

    # 获取施工部位数据
    work_part_value = get_work_part_value(origin_xls, sheet_num)
    # 写入施工部位数据
    write_work_part_value(target_xls,target_xls_sheet_name,work_part_value)

# 数据转换Excel
def convert_excel(origin_file_path, target_xls, log_file):
    # 写记录文件
    log_txt = open(log_file, 'w+')
    # 获取所有在同一个文件夹下的原始文件
    origin_xls_list = get_origin_file_list(origin_file_path)
    # 循环所有的原始文件，保存到一个目标excel
    for origin_xls in origin_xls_list:
        print("正在读取" + origin_xls)
        print("正在读取" + origin_xls, file=log_txt)

        # 打开原始文件
        file_origin_xls = open_xls(origin_xls)
        # 获取原始文件的所有sheet
        file_origin_xls_sheets = file_origin_xls.sheets()
        # 查询sheet数量
        file_origin_xls_sheets_num = len(file_origin_xls_sheets)

        # 打开目标文件
        file_target_xls = open_xls(target_xls)
        # 获取目标文件的所有sheet
        file_target_xls_sheets = file_target_xls.sheets()
        # 查询sheet数量
        file_target_xls_sheets_num = len(file_target_xls_sheets)

        # 如果sheet数量相等执行操作
        if file_origin_xls_sheets_num == file_target_xls_sheets_num :
            # 针对每一个sheet进行数据转换
            for sheet_num in range(0, file_origin_xls_sheets_num):
                print("正在读取" + origin_xls + "的第" + str(sheet_num + 1) + "个标签...", file=log_txt)
                # 写入数据
                target_xls_sheet_name = file_target_xls_sheets[sheet_num].name
                write_to_target_excel(origin_xls,sheet_num,target_xls,target_xls_sheet_name)
    #   关闭
    log_txt.close()
    origin_xls.close()

def start():
    # 源excel文件夹
    origin_file_path = "F:\\mengxiaoqingtest\\xiaoqing"
    # 目标excel
    target_xls = "F:\\结果.xlsx"
    log_file = "F:\\mengxiaoqingtest\\日志.txt"
    convert_excel(origin_file_path, target_xls, log_file)

start()