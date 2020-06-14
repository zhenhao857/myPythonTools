# -*- coding:utf-8 -*-
"""
本工具用于合并多个Excel中的多个sheet（具有相同的表头）至一个Excel中的一个sheet
"""
import os

import xlrd
import xlutils.copy
import xlwings as xw
import xlwt

# C:\Home\Workspace\GitHubWorkSpace\myself\myPythonTools\dependent-package\Lib\site-packages\xlwt\UnicodeUtils.py
# 中
# us = unicode(str(s), encoding)


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


def set_style2(name, height, bold=False):  # 字体设置
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
    borders.left = borders.NO_LINE
    borders.top = borders.NO_LINE
    borders.bottom = borders.MEDIUM
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def set_style3(name, height, bold=False):  # 字体设置
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
    borders.left = borders.MEDIUM
    borders.top = borders.NO_LINE
    borders.bottom = borders.MEDIUM
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
        if os.path.splitext(file_name)[1] == '.xls':
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


# 获取并写入施工里程数据
def get_and_set_construction_mileage_value(row_value, target_xls_sheet):
    work_part_value = row_value[1]
    work_part_value = "DK"+str(work_part_value)[0:3]+"+"+str(work_part_value)[3:len(str(work_part_value))]
    target_xls_sheet.write(4, 3, work_part_value, set_style('宋体', 200, False))
    # # 右边线格式变更
    # target_xls_sheet.write(4, 39, work_part_value, set_style_2('宋体', 200, False))
    # target_xls_sheet.write(5, 39, work_part_value, set_style_2('宋体', 200, False))


# 获取并写入距洞口距离
def get_and_set_opening_distance_value(row_value, target_xls_sheet):
    work_part_value = round(row_value[2], 1)
    target_xls_sheet.write(4, 11, work_part_value, set_style('宋体', 200, False))


# 获取并写入埋深
def get_and_set_about_value(row_value, target_xls_sheet):
    work_part_value = round(row_value[3], 1)
    target_xls_sheet.write(4, 16, work_part_value, set_style('宋体', 200, False))


# 获取并写入水量
def get_and_set_water_volume_value(row_value, target_xls_sheet):
    work_part_value = round(row_value[5], 3)
    work_part_value = "渗（涌）水量："+str(work_part_value)+"，部位拱部及边墙。"
    target_xls_sheet.write(5, 9, work_part_value, set_style('宋体', 200, False))


# 获取并写入实际施工
def get_and_set_actual_construction_value(row_value, target_xls_sheet):
    work_part_value = row_value[6]
    target_xls_sheet.write(6, 6, work_part_value, set_style('宋体', 200, False))


# 获取并写入设计
def get_and_set_design_value(row_value, target_xls_sheet):
    work_part_value = row_value[7]
    target_xls_sheet.write(6, 14, work_part_value, set_style('宋体', 200, False))


# 获取并写入拱顶
def get_and_set_elevation_vault_value(row_value, target_xls_sheet):
    work_part_value = round(row_value[10], 3)
    # work_part_value = round((row_value[1] - 445838) * 3.0 / 1000.0 + 32.76+0.3+8.78+0.55, 3)
    target_xls_sheet.write(9, 6, work_part_value, set_style('宋体', 200, False))


# 获取并写入隧底
def get_and_set_tunnel_bottom_value(row_value, target_xls_sheet):
    work_part_value = round(row_value[11], 3)
    # work_part_value = round((row_value[1] - 445838) * 3.0 / 1000.0 + 32.76+0.3 - 2.98, 3)
    target_xls_sheet.write(9, 14, work_part_value, set_style('宋体', 200, False))


# 获取并写入宽度每侧
def get_and_set_width_each_side_value(row_value, target_xls_sheet):
    work_part_value = round(row_value[12], 1)
    target_xls_sheet.write(10, 6, work_part_value, set_style('宋体', 200, False))


# 获取并写入全宽
def get_and_set_full_width_value(row_value, target_xls_sheet):
    work_part_value = round(row_value[13], 1)
    target_xls_sheet.write(10, 14, work_part_value, set_style('宋体', 200, False))


# 获取并写入最大超挖值
def get_and_set_max_overbreak_value(row_value, target_xls_sheet):
    work_part_value_A = row_value[14]
    work_part_value_B = row_value[15]
    work_part_value_C = row_value[16]
    work_part_value_D = row_value[17]
    work_part_value_E = row_value[18]
    work_part_value_F = row_value[19]
    work_part_value_G = row_value[20]
    work_part_value = row_value[21]
    target_xls_sheet.write(14, 1, work_part_value_A, set_style('宋体', 200, False))
    target_xls_sheet.write(15, 1, work_part_value_B, set_style('宋体', 200, False))
    target_xls_sheet.write(16, 1, work_part_value_C, set_style('宋体', 200, False))
    target_xls_sheet.write(17, 1, work_part_value_D, set_style('宋体', 200, False))
    target_xls_sheet.write(18, 1, work_part_value_E, set_style('宋体', 200, False))
    target_xls_sheet.write(19, 1, work_part_value_F, set_style('宋体', 200, False))
    target_xls_sheet.write(20, 1, work_part_value_G, set_style('宋体', 200, False))
    target_xls_sheet.write(21, 1, work_part_value, set_style('宋体', 200, False))


# 获取并写入实测超挖值
def get_and_set_actual_overbreak_value(row_value, target_xls_sheet):
    work_part_value = row_value[1]
    work_part_value = "DK"+str(work_part_value)[0:3]+"+"+str(work_part_value)[3:len(str(work_part_value))]
    target_xls_sheet.write(13, 5, work_part_value, set_style('宋体', 200, False))

    work_part_value_A = row_value[22]
    work_part_value_B = row_value[23]
    work_part_value_C = row_value[24]
    work_part_value_D = row_value[25]
    work_part_value_E = row_value[26]
    work_part_value_F = row_value[27]
    work_part_value_G = row_value[28]
    work_part_value = row_value[29]
    target_xls_sheet.write(14, 5, work_part_value_A, set_style('宋体', 200, False))
    target_xls_sheet.write(15, 5, work_part_value_B, set_style('宋体', 200, False))
    target_xls_sheet.write(16, 5, work_part_value_C, set_style('宋体', 200, False))
    target_xls_sheet.write(17, 5, work_part_value_D, set_style('宋体', 200, False))
    target_xls_sheet.write(18, 5, work_part_value_E, set_style('宋体', 200, False))
    target_xls_sheet.write(19, 5, work_part_value_F, set_style('宋体', 200, False))
    target_xls_sheet.write(20, 5, work_part_value_G, set_style('宋体', 200, False))
    target_xls_sheet.write(21, 5, work_part_value, set_style('宋体', 200, False))


# 获取并写入日期
def get_and_set_date_value(row_value, target_xls_sheet):
    work_part_value = row_value[8]
    y, m, d = work_part_value.split(".")
    year_month_day = y+" 年 "+m+" 月 "+d+" 日"
    target_xls_sheet.write(27, 0, year_month_day, set_style3('宋体', 200, False))
    target_xls_sheet.write(27, 7, year_month_day, set_style2('宋体', 200, False))
    target_xls_sheet.write(27, 13, year_month_day, set_style2('宋体', 200, False))


# 写入一个sheet
def write_to_target_excel(row_value, target_xls_sheet):
    # 获取并写入施工里程数据
    get_and_set_construction_mileage_value(row_value, target_xls_sheet)

    # 获取并写入距洞口距离
    get_and_set_opening_distance_value(row_value, target_xls_sheet)

    # 获取并写入埋深
    get_and_set_about_value(row_value, target_xls_sheet)

    # 获取并写入水量
    get_and_set_water_volume_value(row_value, target_xls_sheet)

    # 获取并写入实际施工
    get_and_set_actual_construction_value(row_value, target_xls_sheet)

    # 获取并写入设计
    get_and_set_design_value(row_value, target_xls_sheet)

    # 获取并写入拱顶
    get_and_set_elevation_vault_value(row_value, target_xls_sheet)

    # 获取并写入隧底
    get_and_set_tunnel_bottom_value(row_value, target_xls_sheet)

    # 获取并写入宽度每侧
    get_and_set_width_each_side_value(row_value, target_xls_sheet)

    # 获取并写入全宽
    get_and_set_full_width_value(row_value, target_xls_sheet)

    # 获取并写入最大超挖值
    get_and_set_max_overbreak_value(row_value, target_xls_sheet)

    # 获取并写入实测超挖值
    get_and_set_actual_overbreak_value(row_value, target_xls_sheet)

    # 获取并写入日期
    get_and_set_date_value(row_value, target_xls_sheet)


# 插入图片
def insert_iamge(target_xls_sheet, log_txt):
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(target_xls_sheet)
    print("-----------------------------------------------------即将插入图片------------------------------------------------")
    for num in range(0, len(wb.sheets)):
        sht = wb.sheets[num]

        print("在第" + str(num + 1) + "个标签插入图片1")
        print("在第" + str(num + 1) + "个标签插入图片1", file=log_txt)
        image_name1 = os.path.join(os.getcwd(), 'C:\\mengxiaoqing\\20200614\\断面轮廓示意图.png')
        sht.pictures.add(image_name1, left=sht.range('M13').left + 11, top=sht.range('M13').top + 6, width=180,
                         height=138)

        print("在第" + str(num + 1) + "个标签插入图片2")
        print("在第" + str(num + 1) + "个标签插入图片2", file=log_txt)
        image_name2 = os.path.join(os.getcwd(), 'C:\\mengxiaoqing\\20200614\\地质图形素描图.jpg')
        sht.pictures.add(image_name2, left=sht.range('E23').left + 11, top=sht.range('E23').top + 16, width=300,
                         height=150)
    wb.save()
    wb.close()


# 数据转换Excel
def covert_sheet1_taijie(sheet_0, file_target_sheet_num, file_target_xls_cp, log_txt):
    # 需要转换的行数
    # 台阶 1933
    covert_row_num = 1933  # len(sheet_0.rows) - 2
    # 1部 1206
    # 4部 1230
    if covert_row_num > file_target_sheet_num:
        print("需要转换的row多于模板的sheet")
    else:
        # 循环写入
        # for sheet_num in range(0, 3):
        for sheet_num in range(0, covert_row_num):
            # 行号
            row_num = sheet_num+2

            #  获得目标行数据
            row_value = sheet_0.row_values(row_num)

            # 获得目标sheet
            target_xls_sheet = file_target_xls_cp.get_sheet(sheet_num)
            # 设置目标sheet名称
            target_xls_sheet.name = row_num
            print("正在写入sheet" + sheet_0.name + "的第" + str(sheet_num) + "行数据")
            print("正在写入sheet" + sheet_0.name + "的第" + str(sheet_num) + "行数据", file=log_txt)
            write_to_target_excel(row_value, target_xls_sheet)


def convert_excel(origin_xls, target_xls, log_file, save_path_name):

    # 写记录日志文件
    log_txt = open(log_file, 'w+')
    # 打开原始文件
    file_origin_xls = open_xls(origin_xls)
    # 获取原始文件的所有sheet
    file_origin_xls_sheets = file_origin_xls.sheets()
    # 打开目标文件
    file_target_xls = xlrd.open_workbook(target_xls, formatting_info=True)
    # 复制一份用于写入
    file_target_xls_cp = xlutils.copy.copy(file_target_xls)
    # 台阶sheet（sheet0）中的每一行记录 转换 到target 中的 每一个 sheet
    name = file_origin_xls_sheets[0].name
    if name == '台阶':
        print("正在写入台阶sheet" + origin_xls)
        print("正在写入台阶sheet" + origin_xls, file=log_txt)
        covert_sheet1_taijie(file_origin_xls_sheets[0], len(file_target_xls.sheets()), file_target_xls_cp, log_txt)

    file_target_xls_cp.save(save_path_name)

    # 插入图片
    # insert_iamge(save_path_name, log_txt)

    # 关闭
    log_txt.close()


def start():
    # 源excel文件夹
    origin_file = "C:\\mengxiaoqing\\20200614\\原始数据\\新大力寺隧道开挖施工台账.xls"
    # 目标excel
    target_xls = "C:\\mengxiaoqing\\20200614\\台阶-模板.xls"
    log_file = "C:\\mengxiaoqing\\20200614\\台阶日志.txt"
    # 保存的文件名
    save_path_name = "C:\\mengxiaoqing\\20200614\\台阶结果.xls"

    # 执行
    # convert_excel(origin_file, target_xls, log_file, save_path_name)
    # 插入图片
    # 写记录日志文件
    log_txt = open(log_file, 'w+')
    insert_iamge(save_path_name, log_txt)


start()

