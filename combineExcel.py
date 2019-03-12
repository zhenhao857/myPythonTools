# -*- coding:utf-8 -*-

import xlrd,xlsxwriter,os

#源excel文件夹
file_path="F:\\combine1"
#目标excel
end_xls="F:\\合并后3.xlsx"

# 源Excel
def get_source_file_list(file_path):
    allxls = []
    f_list = os.listdir(file_path)
    for fileNAME in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(fileNAME)[1] == '.xlsx':
            fileNAME=file_path+'\\'+fileNAME
            allxls.append(fileNAME)
    return allxls

# 初始化一个目标Excel
def init_excel(destination_file_name):
    #   定义一个目标excel
    endxls=xlsxwriter.Workbook(destination_file_name)
    #   添加一个sheet
    endxls.add_worksheet('sheet1')
    return endxls

# 打开一个源Excel文件
def open_xls(file):
    try:
        fh=xlrd.open_workbook(file)
        return fh
    except Exception as e:
        print("打开文件错误："+e)

#根据excel名以及第几个标签信息就可以得到具体标签的内容
def get_file_value(filename,sheetnum):
    rvalue=[]
    fh=open_xls(filename)
    sheet=fh.sheets()[sheetnum]
    row_num=sheet.nrows
    for rownum in range(0,row_num):
        rvalue.append(sheet.row_values(rownum))
    return rvalue

#将列表file_sheet_value的内容写入目标excel
def write_to_end_excel(file_name,endxls,sheet_value,num_sheet,num_row):
    #   打开目标Excel的sheet
    end_xls_sheet=endxls.get_worksheet_by_name('sheet1')
    num=num_sheet
    num1=num_row
    for row_info in sheet_value:
        num1+=1
        num2=-1
        for colum_info in row_info:
            num2+=1
            #print(num,num1,num2,sheet3)
            #在第num1行的第num2列写入sheet3的内容
            end_xls_sheet.write(num1,num2,colum_info)
            end_xls_sheet.write(num1,num2+1,file_name)
    return num,num1

# 合并Excel
def combine_excel(file_path,end_xls):
    # 写记录文件
    log_txt=open("F:\combine\log2.txt",'w+') 
    
    allxls=get_source_file_list(file_path)
    #   初始化一个目标excel
    endxls=init_excel(end_xls)
    num_sheet=-1
    num_row=-1
    #   循环所有的原始文件
    for file_name in allxls:
        # 保存一个原始文件中的所有列表内容
        # file_all_sheet_value=[]
        print("正在读取"+file_name)
        print("正在读取"+file_name,file=log_txt)
        file_fh=open_xls(file_name)
        file_sheet=file_fh.sheets()
        file_sheet_num=len(file_sheet)
        for sheet_num in range(0,file_sheet_num):
            print("正在读取"+file_name+"的第"+str(sheet_num+1)+"个标签...",file=log_txt)
            sheet_value=get_file_value(file_name,sheet_num)
            # 写入
            num_sheet,num_row=write_to_end_excel(file_name,endxls,sheet_value,num_sheet,num_row)
    #   关闭
    log_txt.close()
    endxls.close()

combine_excel(file_path,end_xls)