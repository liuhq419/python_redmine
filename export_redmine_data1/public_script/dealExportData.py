# coding : utf-8
import pandas as pd
import numpy as np
from excelOpt import *





#设置表格是格式,必须在xlsx中才可以设置表格的格式
def set_format_excel(writer,sheetName,count):
    workbook=writer.book
    worksheet=writer.sheets[sheetName]
    # worksheet.set_column('A:A', 20)
    fmt = workbook.add_format({ 'align': 'left'})
    # fmt1 = workbook.add_format({ 'bold': False})
    #通过遍历 逐行设置行高
    for row in range(count):
        worksheet.set_row(row,18)
    worksheet.set_column('B:B', 37, fmt)
    if sheetName=='Summary':
        worksheet.set_column('C:Z',4)
    else:
        worksheet.set_column('C:Z',15)
    # worksheet.conditional_format('B2:B8', {'type': '3_color_scale'})



# 汇总人员耗时表，添加透视表
def summaryOfLaborHours(excelNames):
    for excelName in excelNames:
        # read excel
        sheet_df = pd.read_excel(excelName, 'Sheet1')
        sheet_df=sheet_df.fillna('')
        # 对数据表进行添加透视表
        sheet_pivoted_df = pd.pivot_table(sheet_df, values=['耗时'], index=['用户', '项目'], columns=['活动', '难易度'],
                                          aggfunc=[np.sum], fill_value='', margins=True, margins_name='总计')
        tCount=len(sheet_df)+5
        sCount=len(sheet_pivoted_df)+20
        # 将读取的数据写入表格
        writer = pd.ExcelWriter(excelName)  # 找到需要写入的excel表格
        sheet_df.to_excel(writer, 'timeLog')  # 将数据写入到sheet1中
        sheet_pivoted_df.to_excel(writer, 'Summary')  # 将数据写入到sheet2中   添加几个sheet,需要重新写几次
        set_format_excel(writer,'timeLog',tCount )  #设置单元格的格式
        set_format_excel(writer,'Summary',sCount )  #设置单元格的格式
        writer.save()  # 保存数据

    # print('人员耗时汇总表导出完毕')



#
# def exportNoSubjectMan(excelNames):
#     for excelName in excelNames:
#         sheet_df=pd.read_excel(excelName,'timelog')
#
#         writer=pd.ExcelWriter(excelName)
#         sheet_df.to_excel(writer,'timelog')
#         sheet_df.query('主题==[""]')
#         sheet_df.to_excel(writer,'noSubjectMan')
#         writer.save()

