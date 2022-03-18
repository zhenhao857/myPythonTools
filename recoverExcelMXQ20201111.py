# -*- coding:utf-8 -*-
"""
本工具用于合并多个Excel中的多个sheet（具有相同的表头）至一个Excel中的一个sheet
"""
import math

import xlrd, xlsxwriter, os, xlutils.copy, xlwt
import xlwings as xw
import image
from datetime import datetime

from xlrd import xldate_as_tuple

from recoverExcelMXQ import get_and_set_make_date_value

from decimal import *

def set_style(name, height, bold=False):  # 字体设置
    """
    设置单元格样式
    :param name: 字体名字
    :param height: 字体大小
    :param bold: 是否加粗
    :return: 返回样式
    """
    style = xlwt.XFStyle()
    font = xlwt.Font()
    alignment = xlwt.Alignment()
    borders = xlwt.Borders()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    alignment.vert = alignment.VERT_CENTER
    alignment.wrap = alignment.WRAP_AT_RIGHT
    alignment.horz = alignment.HORZ_CENTER
    borders.right = borders.THIN
    borders.left = borders.THIN
    borders.top = borders.THIN
    borders.bottom = borders.THIN
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def set_style_2(name, height, bold=False):  # 字体设置
    """
    设置单元格样式
    :param name: 字体名字
    :param height: 字体大小
    :param bold: 是否加粗
    :return: 返回样式
    """
    style = xlwt.XFStyle()
    font = xlwt.Font()
    alignment = xlwt.Alignment()
    borders = xlwt.Borders()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    alignment.vert = alignment.VERT_CENTER
    alignment.wrap = alignment.WRAP_AT_RIGHT
    alignment.horz = alignment.HORZ_CENTER
    borders.right = borders.MEDIUM
    borders.left = borders.THIN
    borders.top = borders.THIN
    borders.bottom = borders.THIN
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


# 获取多个原始Excel列表
def get_origin_file_list(file_path):
    all_file_lists = []
    f_list = os.listdir(file_path)
    for file_name in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(file_name)[1] == '.xlsx' or os.path.splitext(file_name)[1] == '.xls':
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

# 获取并写入序号
def get_and_set_row_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = row_num
    target_xls_sheet.write(row_num, 0, work_part_value, set_style('宋体', 200, False))

# 获取并写入施工部位数据
def get_and_set_work_part_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(5)[2]
    target_xls_sheet.write(row_num, 1, work_part_value, set_style('宋体', 200, False))

# 获取并写入设计强度等级
def get_and_set_strong_num_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(10)[0]
    target_xls_sheet.write(row_num, 2, work_part_value, set_style('宋体', 200, False))

# 获取并写入制件日期
def get_and_set_make_date_value(origin_xls_sheet, target_xls_sheet, row_num):
    cell_type = origin_xls_sheet.cell(12, 0).ctype  # 表格的数据类型
    cell_value = origin_xls_sheet.cell_value(12, 0)
    if cell_type == 2 and cell_value % 1 == 0:  # 如果是整形
        cell = int(cell_value)
        make_date_value = cell
    elif cell_type == 3:
        # 转成datetime对象
        date = datetime(*xldate_as_tuple(cell_value, 0))
        make_date_value = date.strftime('%Y-%m-%d')
    elif cell_type == 1:
        make_date_value = cell_value
    target_xls_sheet.write(row_num, 3, make_date_value, set_style('宋体', 200, False))

# 获取并写入配合比报告编号
def get_and_set_make_report_id_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(10)[5]
    target_xls_sheet.write(row_num, 4, work_part_value, set_style('宋体', 200, False))

# 获取并写入水泥报告编号
def get_and_set_shuini_report_num_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(19)[10]
    target_xls_sheet.write(row_num, 5, work_part_value, set_style('宋体', 200, False))

# 获取并写入掺和料1报告编号
def get_and_set_part1_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(20)[10]
    target_xls_sheet.write(row_num, 6, work_part_value, set_style('宋体', 200, False))

# 获取并写入掺和料2报告编号
def get_and_set_part2_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(21)[10]
    target_xls_sheet.write(row_num, 7, work_part_value, set_style('宋体', 200, False))

# 获取并写入细骨料报告编号
def get_and_set_fine_aggregate_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(22)[10]
    target_xls_sheet.write(row_num, 8, work_part_value, set_style('宋体', 200, False))

# 获取并写入粗骨料报告编号
def get_and_set_coarse_aggregate_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(23)[10]
    target_xls_sheet.write(row_num, 9, work_part_value, set_style('宋体', 200, False))

# 获取并写入外加剂1报告编号
def get_and_set_admixture1_report_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(24)[10]
    target_xls_sheet.write(row_num, 10, work_part_value, set_style('宋体', 200, False))

# 获取并写入外加剂2报告编号
def get_and_set_admixture2_report_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(25)[10]
    target_xls_sheet.write(row_num, 11, work_part_value, set_style('宋体', 200, False))

# 获取并写入拌和水报告编号
def get_and_set_mix_water_report_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(26)[10]
    target_xls_sheet.write(row_num, 12, work_part_value, set_style('宋体', 200, False))

# 获取并写入试件编号
def get_and_set_test_report_value(origin_xls_sheet, target_xls_sheet, row_num):
    work_part_value = origin_xls_sheet.row_values(30)[0]
    target_xls_sheet.write(row_num, 13, work_part_value, set_style('宋体', 200, False))

# 获取并写入获取并写入抗压强度fcu（MPa）
def get_and_set_compressive_strength_num_value(origin_xls_sheet, target_xls_sheet, row_num):

    work_part_value = origin_xls_sheet.row_values(30)[14]
    if work_part_value is '':
        work_part_value = origin_xls_sheet.row_values(29)[14]
        target_xls_sheet.write(row_num, 14, round(Decimal(work_part_value), 1), set_style('宋体', 200, False))
    else:
        target_xls_sheet.write(row_num, 14, round(Decimal(work_part_value), 1), set_style('宋体', 200, False))

    work_part_value = origin_xls_sheet.row_values(33)[14]
    if work_part_value is '':
        work_part_value = origin_xls_sheet.row_values(32)[14]
        target_xls_sheet.write(row_num, 15, round(Decimal(work_part_value), 1), set_style('宋体', 200, False))
    elif work_part_value is '/':
        target_xls_sheet.write(row_num, 15, "/", set_style('宋体', 200, False))
    else:
        target_xls_sheet.write(row_num, 15, round(Decimal(work_part_value), 1), set_style('宋体', 200, False))

def write_to_target_excel(origin_xls_sheet, target_xls_sheet, row_num):

    # 获取并写入序号
    get_and_set_row_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入施工部位数据
    get_and_set_work_part_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入设计强度等级
    get_and_set_strong_num_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入制件日期
    get_and_set_make_date_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入配合比报告编号
    get_and_set_make_report_id_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入水泥报告编号
    get_and_set_shuini_report_num_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入掺和料1报告编号
    get_and_set_part1_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入掺和料2报告编号
    get_and_set_part2_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入细骨料报告编号
    get_and_set_fine_aggregate_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入粗骨料报告编号
    get_and_set_coarse_aggregate_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入外加剂1报告编号
    get_and_set_admixture1_report_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入外加剂2报告编号
    get_and_set_admixture2_report_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入拌和水报告编号
    get_and_set_mix_water_report_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入试件编号
    get_and_set_test_report_value(origin_xls_sheet, target_xls_sheet, row_num)

    # 获取并写入抗压强度fcu（MPa）
    get_and_set_compressive_strength_num_value(origin_xls_sheet, target_xls_sheet, row_num)


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
        # file_origin_xls_sheets_num = 1  # 此段用于测试

        # 打开目标文件
        file_target_xls = xlrd.open_workbook(target_xls, formatting_info=True)  # 打开文件

        # 复制一份用于写入
        file_target_xls_cp = xlutils.copy.copy(file_target_xls)

        # 如果sheet数量适应则执行操作
        if 1 == 1:  # file_origin_xls_sheets_num == 1611:
            # 针对每一个sheet进行数据转换
            for sheet_num in range(0, file_origin_xls_sheets_num):
                print("正在读取" + origin_xls + "的第" + str(sheet_num + 1) + "个标签...")
                print("正在读取" + origin_xls + "的第" + str(sheet_num + 1) + "个标签...", file=log_txt)
                # 获取原始文件的sheet
                origin_xls_sheet = file_origin_xls.sheets()[sheet_num]
                # 原始文件的sheet名称
                origin_xls_sheet_name = origin_xls_sheet.name
                # 打开目标Excel的sheet
                target_xls_sheet = file_target_xls_cp.get_sheet(0)  # 写第一个sheet

                # 写入数据
                print("正在写入" + origin_xls + "的" + origin_xls_sheet_name + ": 第" + str(sheet_num + 1) + "个标签...")
                print("正在写入" + origin_xls + "的" + origin_xls_sheet_name + ": 第" + str(sheet_num + 1) + "个标签...", file=log_txt)

                # 确定写入行
                if sheet_num+1 == 901:
                    write_to_target_excel(origin_xls_sheet, target_xls_sheet, sheet_num + 1)
                else:
                    write_to_target_excel(origin_xls_sheet, target_xls_sheet, sheet_num+1)

        file_target_xls_cp.save(os.path.splitext(origin_xls)[0]+"_结果.xls")

    log_txt.close()


def start_20201112():
    # 源excel文件夹
    origin_file_path = "C:\\mengxiaoqing\\20201112\\原始数据2"
    # 目标excel
    target_xls = "C:\\mengxiaoqing\\20201112\\预制厂混凝土浇筑台账.xls"
    log_file = "C:\\mengxiaoqing\\20201112\\日志.txt"
    convert_excel(origin_file_path, target_xls, log_file)


start_20201112()
# print((math.modf(3.45*10.0))[1])
# print(round(2.4))
# print(round(Decimal(str(2.45)), 1))
# print(round(Decimal('2.45'), 1))
# print(round(2.6))
# print(round(3.4))
# print(round(3.5))
# print(round(3.6))
