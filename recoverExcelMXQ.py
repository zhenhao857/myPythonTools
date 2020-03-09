# -*- coding:utf-8 -*-
"""
本工具用于合并多个Excel中的多个sheet（具有相同的表头）至一个Excel中的一个sheet
"""
import xlrd, xlsxwriter, os, xlutils.copy, xlwt
import xlwings as xw
import image
from datetime import datetime

from xlrd import xldate_as_tuple

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


# 获取并写入施工部位数据
def get_and_set_work_part_value(origin_xls_sheet, target_xls_sheet):
    work_part_value = origin_xls_sheet.row_values(5)[2]
    target_xls_sheet.write(4, 30, work_part_value, set_style('宋体', 200, False))
    # 右边线格式变更
    target_xls_sheet.write(4, 39, work_part_value, set_style_2('宋体', 200, False))
    target_xls_sheet.write(5, 39, work_part_value, set_style_2('宋体', 200, False))


# 获取并写入代表数量
def get_and_set_deputy_num_value(origin_xls_sheet, target_xls_sheet):
    deputy_num_value = origin_xls_sheet.row_values(6)[2]
    target_xls_sheet.write(4, 41, "代表数量", set_style('宋体', 200, False))
    target_xls_sheet.write(5, 41, deputy_num_value, set_style('宋体', 200, False))


# 获取并写入制件日期
def get_and_set_make_date_value(origin_xls_sheet, target_xls_sheet):
    cell_type = origin_xls_sheet.cell(12, 0).ctype  # 表格的数据类型
    cell_value = origin_xls_sheet.cell_value(12, 0)
    if cell_type == 2 and cell_value % 1 == 0:  # 如果是整形
        cell = int(cell_value)
        make_date_value = cell
    elif cell_type == 3:
        # 转成datetime对象
        date = datetime(*xldate_as_tuple(cell_value, 0))
        make_date_value = date.strftime('%Y.%m.%d')
    elif cell_type == 1:
        make_date_value = cell_value

    target_xls_sheet.write(17, 41, "制件日期", set_style('宋体', 200, False))
    target_xls_sheet.write(18, 41, make_date_value, set_style('宋体', 200, False))


# 获取并写入型号产地
def get_and_set_make_place_value(origin_xls_sheet, target_xls_sheet):
    # cement
    # Fine aggregate
    # Coarse aggregate
    # Additive1
    # Additive2
    # Mixing water

    # 水泥
    cement_value = origin_xls_sheet.row_values(19)[4]
    target_xls_sheet.write(7, 7, cement_value, set_style('宋体', 200, False))
    # 细骨料
    fine_aggregate_value = origin_xls_sheet.row_values(20)[4]
    target_xls_sheet.write(8, 7, fine_aggregate_value, set_style('宋体', 200, False))
    # 粗骨料
    coarse_aggregate_value = origin_xls_sheet.row_values(21)[4]
    target_xls_sheet.write(9, 7, coarse_aggregate_value, set_style('宋体', 200, False))
    # 外加剂1
    additive1_value = origin_xls_sheet.row_values(22)[4]
    target_xls_sheet.write(10, 7, additive1_value, set_style('宋体', 200, False))
    # 外加剂2
    additive2_value = origin_xls_sheet.row_values(23)[4]
    target_xls_sheet.write(11, 7, additive2_value, set_style('宋体', 200, False))
    # 拌和水
    mixing_water_value = origin_xls_sheet.row_values(24)[4]
    target_xls_sheet.write(12, 7, mixing_water_value, set_style('宋体', 200, False))


# 获取并写入试验报告编号和结论
def get_and_set_report_num_value(origin_xls_sheet, target_xls_sheet):
    # 水泥
    report_num_cement_value = origin_xls_sheet.row_values(19)[10]
    target_xls_sheet.write(7, 23, report_num_cement_value, set_style('宋体', 200, False))
    # 细骨料
    report_num_fine_aggregate_value = origin_xls_sheet.row_values(20)[10]
    target_xls_sheet.write(8, 23, report_num_fine_aggregate_value, set_style('宋体', 200, False))
    # 粗骨料
    report_num_coarse_aggregate_value = origin_xls_sheet.row_values(21)[10]
    target_xls_sheet.write(9, 23, report_num_coarse_aggregate_value, set_style('宋体', 200, False))
    # 外加剂1
    report_num_additive1_value = origin_xls_sheet.row_values(22)[10]
    target_xls_sheet.write(10, 23, report_num_additive1_value, set_style('宋体', 200, False))
    # 外加剂2
    report_num_additive2_value = origin_xls_sheet.row_values(23)[10]
    target_xls_sheet.write(11, 23, report_num_additive2_value, set_style('宋体', 200, False))
    # 拌和水
    report_num_mixing_water_value = origin_xls_sheet.row_values(24)[10]
    target_xls_sheet.write(12, 23, report_num_mixing_water_value, set_style('宋体', 200, False))


# 写入一个sheet
def write_to_target_excel(origin_xls_sheet, target_xls_sheet):
    # 获取并写入施工部位数据
    get_and_set_work_part_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入代表数量
    get_and_set_deputy_num_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入制件日期
    get_and_set_make_date_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入型号产地
    get_and_set_make_place_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入试验报告编号和结论
    get_and_set_report_num_value(origin_xls_sheet, target_xls_sheet)


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
        file_target_xls = xlrd.open_workbook(target_xls, formatting_info=True)  # 打开文件
        # 获取目标文件的所有sheet
        file_target_xls_sheets = file_target_xls.sheets()
        # 查询sheet数量
        file_target_xls_sheets_num = len(file_target_xls_sheets)

        # 复制一份用于写入
        file_target_xls_cp = xlutils.copy.copy(file_target_xls)

        # 如果sheet数量适应则执行操作
        if file_origin_xls_sheets_num <= file_target_xls_sheets_num:
            # 针对每一个sheet进行数据转换
            for sheet_num in range(0, file_origin_xls_sheets_num):
                print("正在读取" + origin_xls + "的第" + str(sheet_num + 1) + "个标签...", file=log_txt)
                # 获取原始文件的sheet
                origin_xls_sheet = file_origin_xls.sheets()[sheet_num]
                # 原始文件的sheet名称
                origin_xls_sheet_name = origin_xls_sheet.name
                # 打开目标Excel的sheet
                target_xls_sheet = file_target_xls_cp.get_sheet(sheet_num)  # target_xls_sheet_name
                # 设置目标sheet名称
                target_xls_sheet.name = origin_xls_sheet_name

                # 写入数据
                print("正在写入" + origin_xls + "的" + origin_xls_sheet_name + ": 第" + str(sheet_num + 1) + "个标签...", file=log_txt)
                write_to_target_excel(origin_xls_sheet, target_xls_sheet)

    file_target_xls_cp.save("C:\\mengxiaoqing\\test1\\123.xls")

    # 插入图片
    image_name = os.path.join(os.getcwd(), 'C:\\mengxiaoqing\\test1\\1.png')
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open('C:\\mengxiaoqing\\test1\\123.xls')
    print("-----------------------------------------------------即将插入图片------------------------------------------------")
    for num in range(0, len(wb.sheets)):
        sht = wb.sheets[num]
        print("正在插入图片的第" + str(num + 1) + "个标签...", file=log_txt)
        sht.pictures.add(image_name, left=sht.range('H21').left+11, top=sht.range('H21').top+6, width=180, height=138)
    wb.save()
    wb.close()
    # 关闭
    log_txt.close()

def start():
    # 源excel文件夹
    origin_file_path = "C:\\mengxiaoqing\\test1\\mengxiaoqing"
    # 目标excel
    target_xls = "C:\\mengxiaoqing\\test1\\结果.xls"
    log_file = "C:\\mengxiaoqing\\test1\\日志.txt"
    convert_excel(origin_file_path, target_xls, log_file)

start()
