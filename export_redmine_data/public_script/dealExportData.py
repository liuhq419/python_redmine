# coding : utf-8
import pandas as pd
import numpy as np

# 暂时不运行
# 汇总人员耗时表，添加透视表
def summaryOfLaborHours(excelNames):
    for excelName in excelNames:
        # read excel
        sheet_df = pd.read_excel(excelName, 'Sheet1')
        # 对数据表进行添加透视表
        sheet_pivoted_df = pd.pivot_table(sheet_df, values=['耗时'], index=['用户', '项目'], columns=['活动', '难易度'],
                                          aggfunc=np.sum, fill_value='', margins=True, margins_name='总计')
        # 将读取的数据写入表格
        writer = pd.ExcelWriter(excelName)  # 找到需要写入的excel表格
        sheet_df.to_excel(writer, 'timeLog')  # 将数据写入到sheet1中
        sheet_pivoted_df.to_excel(writer, 'Sheet1')  # 将数据写入到sheet2中   添加几个sheet,需要重新写几次
        writer.save()  # 保存数据
    print('人员耗时汇总表导出完毕')


def exportNoSubjectMan(excelNames):
    for excelName in excelNames:
        sheet_df=pd.read_excel(excelName,'timelog')

        writer=pd.ExcelWriter(excelName)
        sheet_df.to_excel(writer,'timelog')
        sheet_df.query('主题==[""]')
        sheet_df.to_excel(writer,'noSubjectMan')
        writer.save()

