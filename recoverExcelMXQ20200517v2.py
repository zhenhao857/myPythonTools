# -*- coding:utf-8 -*-
"""
本工具用于合并多个Excel中的多个sheet（具有相同的表头）至一个Excel中的一个sheet
"""
import xlrd, os, xlutils.copy, xlwt


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
    borders.right = borders.THIN
    borders.left = borders.MEDIUM
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


# 写入一个sheet
def write_to_target_excel(origin_xls_sheet, target_xls_sheet):

    # 20200517日修订
    # 获取并写入施工部位数据
    get_and_set_work_part_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入使用强度数据
    get_and_set_strength_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入报告编号数据
    get_and_set_report_num_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入记录编号数据
    get_and_set_record_num_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入评定编号数据
    get_and_set_evaluate_num_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入报告日期数据
    get_and_set_report_date_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入配合比例数据
    get_and_set_fit_ratio_num_value(origin_xls_sheet, target_xls_sheet)

    # 获取并写入抗压强度数据，包含第二页处理
    get_and_set_compressive_strength_num_value(origin_xls_sheet, target_xls_sheet)

    # 变更列宽
    change_col_width2(origin_xls_sheet, target_xls_sheet)


# 获取并写入施工部位数据,2处
def get_and_set_work_part_value(origin_xls_sheet, target_xls_sheet):
    # with1 = target_xls_sheet.col(0).width
    # print(with1)
    work_part_value = origin_xls_sheet.row_values(6)[2]
    target_xls_sheet.write(5, 2, work_part_value, set_style('宋体', 200, False))


# 获取并写入使用强度数据
def get_and_set_strength_value(origin_xls_sheet, target_xls_sheet):
    work_part_value = origin_xls_sheet.row_values(11)[0]
    target_xls_sheet.write(6, 2, work_part_value, set_style('宋体', 200, False))

    # 写入第二页
    work_part_value_real = work_part_value[1:len(work_part_value)]
    target_xls_sheet.write(19, 48, work_part_value_real, set_style('宋体', 200, False))


# 获取并写入报告编号数据
def get_and_set_report_num_value(origin_xls_sheet, target_xls_sheet):
    work_part_value = origin_xls_sheet.row_values(4)[13]+origin_xls_sheet.row_values(4)[19]
    target_xls_sheet.write(3, 10, work_part_value.strip(), set_style('宋体', 200, False))


# 获取并写入记录编号数据
def get_and_set_record_num_value(origin_xls_sheet, target_xls_sheet):
    work_part_value = origin_xls_sheet.row_values(4)[13]+origin_xls_sheet.row_values(4)[19]
    target_xls_sheet.write(4, 10, work_part_value.strip(), set_style('宋体', 200, False))


# 获取并写入评定编号数据
def get_and_set_evaluate_num_value(origin_xls_sheet, target_xls_sheet):
    work_part_value = origin_xls_sheet.row_values(4)[13]+origin_xls_sheet.row_values(4)[19]
    work_part_value = work_part_value.strip()
    # 处理work_part_value，替换KYPD
    work_part_value = work_part_value[0:13]+'KYPD'+work_part_value[17:len(work_part_value)]
    target_xls_sheet.write(5, 10, work_part_value, set_style('宋体', 200, False))


# 获取并写入报告日期数据
def get_and_set_report_date_value(origin_xls_sheet, target_xls_sheet):
    work_part_value = origin_xls_sheet.row_values(7)[13]
    target_xls_sheet.write(6, 10, work_part_value, set_style('宋体', 200, False))


# 获取并写入配合比例数据
def get_and_set_fit_ratio_num_value(origin_xls_sheet, target_xls_sheet):
    work_part_value = origin_xls_sheet.row_values(11)[5]
    target_xls_sheet.write(9, 8, work_part_value, set_style('宋体', 200, False))


# 获取并写入抗压强度数据
def get_and_set_compressive_strength_num_value(origin_xls_sheet, target_xls_sheet):
    list_compressive_strength = []

    work_part_value_1 = origin_xls_sheet.row_values(31)[21]
    if work_part_value_1 != '':
        work_part_value_1 = float(work_part_value_1)
        # 写入第一页抗压轻度
        target_xls_sheet.write(12, 5, work_part_value_1, set_style('宋体', 200, False))
        # 获取第一页报告编号
        report_num_value = origin_xls_sheet.row_values(4)[13] + origin_xls_sheet.row_values(4)[19]
        # 写入序号与第一页报告编号
        target_xls_sheet.write(12, 2, report_num_value, set_style('宋体', 200, False))
        target_xls_sheet.write(12, 0, '1', set_style_2('宋体', 200, False))
        # 写入第二页报告编号,左边线格式变更
        target_xls_sheet.write(12, 17, report_num_value, set_style_2('宋体', 200, False))
        # 写入第二页抗压强度
        target_xls_sheet.write(12, 22, work_part_value_1, set_style('宋体', 200, False))

        list_compressive_strength.append(work_part_value_1)

    work_part_value_2 = origin_xls_sheet.row_values(34)[21]
    if work_part_value_2 != '':
        work_part_value_2 = float(work_part_value_2)
        # 写入第一页抗压轻度
        target_xls_sheet.write(13, 5, work_part_value_2, set_style('宋体', 200, False))
        # 获取第一页报告编号
        report_num_value = origin_xls_sheet.row_values(4)[13] + origin_xls_sheet.row_values(4)[19]
        # 写入序号与第一页报告编号
        target_xls_sheet.write(13, 2, report_num_value, set_style('宋体', 200, False))
        target_xls_sheet.write(13, 0, '2', set_style_2('宋体', 200, False))
        # 写入第二页报告编号
        target_xls_sheet.write(13, 17, report_num_value, set_style_2('宋体', 200, False))
        # 写入第二页抗压强度
        target_xls_sheet.write(13, 22, work_part_value_2, set_style('宋体', 200, False))

        list_compressive_strength.append(work_part_value_2)

    work_part_value_3 = origin_xls_sheet.row_values(37)[21]
    if work_part_value_3 != '':
        work_part_value_3 = float(work_part_value_3)
        # 写入第一页抗压轻度
        target_xls_sheet.write(14, 5, work_part_value_3, set_style('宋体', 200, False))
        # 获取第一页报告编号
        report_num_value = origin_xls_sheet.row_values(4)[13] + origin_xls_sheet.row_values(4)[19]
        # 写入序号与第一页报告编号
        target_xls_sheet.write(14, 2, report_num_value, set_style('宋体', 200, False))
        target_xls_sheet.write(14, 0, '3', set_style_2('宋体', 200, False))
        # 写入第二页报告编号
        target_xls_sheet.write(14, 17, report_num_value, set_style_2('宋体', 200, False))
        # 写入第二页抗压强度
        target_xls_sheet.write(14, 22, work_part_value_3, set_style('宋体', 200, False))

        list_compressive_strength.append(work_part_value_3)

    sum_compressive_strength = 0.0
    min_compressive_strength = list_compressive_strength[0]
    for i in list_compressive_strength:
        sum_compressive_strength = sum_compressive_strength + i
        if min_compressive_strength > i:
            min_compressive_strength = i
    target_xls_sheet.write(12, 31, round(sum_compressive_strength/len(list_compressive_strength), 1), set_style('宋体', 200, False))
    target_xls_sheet.write(12, 38, min_compressive_strength, set_style('宋体', 200, False))


# 修订Excel列宽
def change_col_width2(origin_xls_sheet, target_xls_sheet):
    width_39 = 1500   # 3.89
    width_53 = 1800  # 1560   5.33
    width_75 = 2115  # 2110   7.44
    width_44 = 1350  # 4.44
    width_24 = 820    #800
    width_52 = 1530   # 5.22
    width_41 = 1300   # 1330
    width_37 = 1170   # 3.78
    width_49 = 1450   # 4.67 1400
    width_22 = 900
    width_041 = 133
    width_018 = 60
    width_23 = 950
    width_40 = 1250

    target_xls_sheet.col(1).width = 1800    # b
    target_xls_sheet.col(4).width = 1800    # e
    target_xls_sheet.col(7).width = 1800    # h
    target_xls_sheet.col(9).width = 1800    # j
    target_xls_sheet.col(12).width = 1800    # m
    target_xls_sheet.col(15).width = 1800    # p

    target_xls_sheet.col(16).width = width_39    # q
    target_xls_sheet.col(17).width = width_22    # r
    target_xls_sheet.col(18).width = width_22    # s
    target_xls_sheet.col(19).width = width_22    # t
    target_xls_sheet.col(20).width = width_22    # u
    target_xls_sheet.col(21).width = width_22    # v
    target_xls_sheet.col(22).width = width_22    # w
    target_xls_sheet.col(23).width = width_22    # x
    target_xls_sheet.col(24).width = width_22    # y
    target_xls_sheet.col(25).width = width_22    # z
    target_xls_sheet.col(26).width = width_22    # aa
    target_xls_sheet.col(27).width = width_22    # ab
    target_xls_sheet.col(28).width = width_22    # ac
    target_xls_sheet.col(29).width = width_22    # ad
    target_xls_sheet.col(30).width = width_22    # ae
    target_xls_sheet.col(31).width = width_22    # af
    target_xls_sheet.col(32).width = width_22    # ag
    target_xls_sheet.col(33).width = width_22    # ah
    target_xls_sheet.col(34).width = width_22    # ai
    target_xls_sheet.col(35).width = width_22    # aj
    target_xls_sheet.col(35).width = width_22    # ak
    target_xls_sheet.col(37).width = width_22    # al
    target_xls_sheet.col(38).width = width_22    # am
    target_xls_sheet.col(39).width = width_22    # an
    target_xls_sheet.col(40).width = width_22    # ao
    target_xls_sheet.col(41).width = width_22    # ap
    target_xls_sheet.col(42).width = width_22    # aq
    target_xls_sheet.col(43).width = width_041    # ar
    target_xls_sheet.col(44).width = width_018    # as
    target_xls_sheet.col(45).width = width_23    # at
    target_xls_sheet.col(46).width = width_23    # au
    target_xls_sheet.col(47).width = width_23    # av
    target_xls_sheet.col(48).width = width_40    # aw

# 修订Excel列宽
def change_col_width(origin_xls_sheet, target_xls_sheet):
    width_39 = 1500   # 3.89
    width_53 = 1700  # 1560   5.33
    width_75 = 2115  # 2110   7.44
    width_44 = 1350  # 4.44
    width_24 = 820    #800
    width_52 = 1530   # 5.22
    width_41 = 1300   # 1330
    width_37 = 1170   # 3.78
    width_49 = 1450   # 4.67 1400
    width_22 = 900
    width_041 = 133
    width_018 = 60
    width_23 = 950
    width_40 = 1250
    target_xls_sheet.col(0).width = width_39    # a
    target_xls_sheet.col(1).width = width_39    # b
    target_xls_sheet.col(2).width = width_39    # c
    target_xls_sheet.col(3).width = width_53    # d
    target_xls_sheet.col(4).width = width_53    # e
    target_xls_sheet.col(5).width = width_75    # f
    target_xls_sheet.col(6).width = width_44    # g
    target_xls_sheet.col(7).width = width_24    # h
    target_xls_sheet.col(8).width = width_44    # i
    target_xls_sheet.col(9).width = width_52    # j
    target_xls_sheet.col(10).width = width_44    # k
    target_xls_sheet.col(11).width = width_41    # l
    target_xls_sheet.col(12).width = width_24    # m
    target_xls_sheet.col(13).width = width_41    # n
    target_xls_sheet.col(14).width = width_37    # o
    target_xls_sheet.col(15).width = width_49    # p
    target_xls_sheet.col(16).width = width_39    # q
    target_xls_sheet.col(17).width = width_22    # r
    target_xls_sheet.col(18).width = width_22    # s
    target_xls_sheet.col(19).width = width_22    # t
    target_xls_sheet.col(20).width = width_22    # u
    target_xls_sheet.col(21).width = width_22    # v
    target_xls_sheet.col(22).width = width_22    # w
    target_xls_sheet.col(23).width = width_22    # x
    target_xls_sheet.col(24).width = width_22    # y
    target_xls_sheet.col(25).width = width_22    # z
    target_xls_sheet.col(26).width = width_22    # aa
    target_xls_sheet.col(27).width = width_22    # ab
    target_xls_sheet.col(28).width = width_22    # ac
    target_xls_sheet.col(29).width = width_22    # ad
    target_xls_sheet.col(30).width = width_22    # ae
    target_xls_sheet.col(31).width = width_22    # af
    target_xls_sheet.col(32).width = width_22    # ag
    target_xls_sheet.col(33).width = width_22    # ah
    target_xls_sheet.col(34).width = width_22    # ai
    target_xls_sheet.col(35).width = width_22    # aj
    target_xls_sheet.col(35).width = width_22    # ak
    target_xls_sheet.col(37).width = width_22    # al
    target_xls_sheet.col(38).width = width_22    # am
    target_xls_sheet.col(39).width = width_22    # an
    target_xls_sheet.col(40).width = width_22    # ao
    target_xls_sheet.col(41).width = width_22    # ap
    target_xls_sheet.col(42).width = width_22    # aq
    target_xls_sheet.col(43).width = width_041    # ar
    target_xls_sheet.col(44).width = width_018    # as
    target_xls_sheet.col(45).width = width_23    # at
    target_xls_sheet.col(46).width = width_23    # au
    target_xls_sheet.col(47).width = width_23    # av
    target_xls_sheet.col(48).width = width_40    # aw


    # list_col = [3.9, 3.9, 3.9, 5.3, 5.3, 7.5, 4.4, 2.4, 4.4, 5.2, 4.4, 4.1, 2.4, 4.1, 3.7, 4.9, 3.9, 2.2, 2.2, 2.2, 2.2,
    #             2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 2.2, 0.41, 0.18, 2.3, 2.3,
    #             2.3, 4.0]
    # for i in range(0, len(list_col)):
    #     width_origin = round(list_col[i]*2300.0/11.0+384.0, 1)
    #     target_xls_sheet.col(i).width = width_origin


# # 利用模板创建与打开的file_origin_xls相同数量的sheet页,且文件名称相同，前缀"结果"
# def generate_target_file_xls(target_xls_template, file_origin_xls_sheets_num):
#     target_xls_template_open = xlrd.open_workbook(target_xls_template, formatting_info=True)  # 打开文件
#     target_xls_template_copy = xlutils.copy.copy(target_xls_template_open)
#     sheet_0 = target_xls_template_copy.sheets()[0]
#     target_xls_template_copy.add_sheet(sheet_0)
#     target_xls_template_copy.save("C:\\mengxiaoqing\\test1\\结果_"+target_xls_template.name+".xls")
#     return target_xls_template_copy


# 数据转换Excel
def convert_excel(origin_file_path, target_xls, log_file, document_name):

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

        # # 利用模板创建与打开的file_origin_xls相同数量的sheet页,且文件名称相同，前缀"结果"
        # target_xls = generate_target_file_xls(target_xls_template, file_origin_xls_sheets_num)

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
        # if 1 <= 2:
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

        file_target_xls_cp.save("C:\\mengxiaoqing\\test1\\"+document_name+"\\结果_"+document_name+".xls")

    # 关闭
    log_txt.close()


def start():
    # document_name = "北端上行垫石"
    # document_name = "北端上行联络线承台"
    # document_name = "北端上行联络线墩身"
    document_name = "北端下行垫石"
    # document_name = "北端下行联络线承台"
    # document_name = "北端下行联络线墩身"
    # document_name = "南端上行垫石"
    # document_name = "南端上行墩身2014.12-2015.6月份"
    # document_name = "南端上行联络线承台"
    # document_name = "南端下行垫石"
    # document_name = "南端下行联络线承台"
    # document_name = "南端下行联络线墩身"
    # 源excel文件夹
    origin_file_path = "C:\\mengxiaoqing\\test1\\"+document_name+"\\mengxiaoqing"
    # 目标excel
    target_xls = "C:\\mengxiaoqing\\test1\\"+document_name+"\\"+document_name+".xls"
    log_file = "C:\\mengxiaoqing\\test1\\"+document_name+"\\日志.txt"
    convert_excel(origin_file_path, target_xls, log_file, document_name)


start()
