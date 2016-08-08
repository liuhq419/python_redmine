# -*- coding: utf-8 -*-
'''函数功能：创建excel；插入数据到excel'''
import os,sys
import xlrd,xlwt
from xlutils.copy import copy

#设置单元格的样式
def setExcelStyle():
    #设置单元格的对齐方式
    alignment=xlwt.Alignment()
    alignment.horz=xlwt.Alignment.HORZ_CENTER
    alignment.vert=xlwt.Alignment.VERT_CENTER
    #设置边框
    borders=xlwt.Borders()
    borders.left=1
    borders.right=1
    borders.top=1
    borders.bottom=1

    #设置单元格的对齐方式
    alignment1=xlwt.Alignment()
    alignment1.horz=xlwt.Alignment.HORZ_LEFT
    alignment1.vert=xlwt.Alignment.VERT_CENTER

    #font1的格式为表头 加粗 宋体 16号
    font1 = xlwt.Font()
    font1.name = '宋体'
    font1.bold = True
    font1.height = 0x00140   #字体大小为十六进制转为十进制 除以20   0x00104=320
    style1 = xlwt.XFStyle()  # create the style
    style1.font = font1
    style1.alignment=alignment
    style1.borders=borders

    #font2为 宋体 加粗 10号
    font2 = xlwt.Font()
    font2.name = '宋体'
    font2.bold = True
    font2.height = 0x00C8  # 字体大小为十六进制转为十进制 除以20   0x00DC=220
    style2 = xlwt.XFStyle()  # create the style
    style2.font = font2
    style2.alignment=alignment
    style2.borders=borders

    #font3为 宋体  10号
    font3 = xlwt.Font()
    font3.name = '宋体'
    font3.bold = False
    font3.height = 0x00C8  # 字体大小为十六进制转为十进制 除以30   0x00DC=220
    style3 = xlwt.XFStyle()  # create the style
    style3.font = font3
    style3.alignment=alignment  #设置对齐方式
    style3.borders=borders


    style4 = xlwt.XFStyle()  # create the style
    style4.font = font3
    style4.alignment=alignment1  #设置对齐方式
    style4.borders=borders

    return (style1,style2,style3,style4)
#三种不同的格式，公共变量
style1, style2, style3,style4 = setExcelStyle()
#行高
tall_style=xlwt.easyxf('font:height 280;')  #普通行行高
fir_tall_style=xlwt.easyxf('font:height 400;')  #普通行行高


'''write_merge(x, x + m, y, y + n, string, sytle)
x表示行，y表示列，m表示跨行个数，n表示跨列个数，string表示要写入的单元格内容，
style表示单元格样式。其中，x，y，m，n，都是以0开始计算的。'''

'''建立excel，存储查询结果'''
'''该函数在insertIntoExcel()插入数据函数中调用，name,filePath,lists在insertExcel()中获得'''
def createExcel(date,PATH,lists,depart='',str= 'PM系统工时填报'):
    # 新建xlsx，新建名为sheet1的工作簿
    file=xlwt.Workbook()  #encoding='ascii'
    timeSheet=file.add_sheet('timeLog',cell_overwrite_ok=True)  #添加人员维度的工作簿

    timeSheet.row(0).set_style(tall_style)  #设置行高
    for i in range(0,len(lists)):
        if str=='PM系统工时填报':
            timeSheet.col(i).width=256*14
            timeSheet.col(0).width=256*32
            timeSheet.col(4).width=256*32
        else:
            timeSheet.col(i).width=256*10
            timeSheet.col(1).width=256*32
            timeSheet.col(4).width=256*32
            timeSheet.col(6).width=256*20
            timeSheet.col(7).width=256*20
            if depart:
                timeSheet.col(9).width=256*20
            else:
                timeSheet.col(8).width=256*20
                timeSheet.col(10).width=256*20

        timeSheet.row(0).set_style(tall_style)
        timeSheet.write(0,i,lists[i],style2)    # 逐个插入lists列表中的数据，即为表头

    filePath = PATH + '\\exportResult\\'
    excelName= str+'\\'+depart + str + date + '.xlsx'
    pathExcelName = filePath +excelName
    file.save(pathExcelName)  # 保存excel
    return pathExcelName

'''插入数据到excel中
context 为所要插入的数据，'''
def insertIntoExcel(content,PATH,date,depart='',str= 'PM系统工时填报'):
    filePath = PATH + '\\exportResult\\'
    excelName = filePath +str+'\\'+depart +str+ date + '.xlsx'
    #打开要插入的表，并将数据复制到新的表中
    oldExcel=xlrd.open_workbook(excelName, formatting_info=True)
    newExcel=copy(oldExcel)
    sheet=oldExcel.sheet_by_index(0)
    newSheet=newExcel.get_sheet(0)

    #逐行插入数据
    r=sheet.nrows
    for col in range(0,sheet.ncols):
        #table.nrows控制始终将数据插入到当前文件的最后一行
        newSheet.row(r-1).set_style(tall_style)
        newSheet.row(r).set_style(tall_style)
        newSheet.write(r,col,content[col],style4)
    newExcel.save(excelName)


def missingPerson(PATH,xlsNames,newdeparts,staffExcel='格网各部门人员清单.xls'):
    staffxls=PATH+'\\public_script\\'+staffExcel
    for depart in newdeparts:
        staffXls=xlrd.open_workbook(staffxls)
        staffSht=staffXls.sheet_by_name(depart)
        totalStaff=staffSht.col_values(0,start_rowx=1)
        missingStaff=[]
        for xlsName in xlsNames:
            if depart in xlsName:
                timeXls=xlrd.open_workbook(xlsName,formatting_info=True)
                newTimeXLs=copy(timeXls)
                timeSht=timeXls.sheet_by_index(2)
                newTimeSht=newTimeXLs.get_sheet(2)
                newTimeSht1=newTimeXLs.get_sheet(1)
                wroteStaff=timeSht.col_values(0,6,timeSht.nrows-2)
                wroteStaff = str(wroteStaff)
                wroteStaff = wroteStaff.replace('[', '')
                wroteStaff = wroteStaff.replace(']', '')
                wroteStaff = wroteStaff.replace("'", '')
                wroteStaff = wroteStaff.replace(" ", '')
                # print('wroteStaff:',wroteStaff)
                # print('totalStaff:',totalStaff)
                for total in totalStaff:
                    total=str(total).replace(' ','')
                    if total not in wroteStaff:
                        missingStaff.append(total)
                        miss=['本月填报缺失人员:']+missingStaff
                        miss=str(miss)
                        miss=miss.replace('[','')
                        miss=miss.replace(']','')
                        miss=miss.replace("'",'')
                        newTimeSht.write(1,0,miss)
                        newTimeSht1.write(1,0,miss)
                        newTimeXLs.save(xlsName)
            else:continue
# missingPerson()

import pandas as pd
import numpy as np

# 读取excel表格中的数据，并进行相应的处理，本例是添加透视表
def readExcel(excelNames,PATH,newdeparts):
    xlsNames = []
    for excelName in excelNames:
        excel_df = pd.ExcelFile(excelName)#,engine='xlsxwriter')  # 获得excel
        sheet_df = excel_df.parse('timeLog')  # 获得excel的工作簿   #返回dataFrame
        sheet_df = sheet_df.fillna('-')  # 空值处理，将空值用''填充
        try:
            # 对数据表进行添加透视表
            sheet_pivoted_df = pd.pivot_table(sheet_df, values='耗时', index=['用户', '项目'], columns=['活动', '难易度'],
                                              aggfunc={'耗时': np.sum}, fill_value='', margins=True,
                                              margins_name='总计')
            pivoted_df_user = pd.pivot_table(sheet_df, values='耗时', index=['用户'], columns=['活动', '难易度'],
                                             aggfunc={'耗时': np.sum}, fill_value='', margins=True,
                                             margins_name='总计')
            pivoted_df_user.insert(0, '活动', '汇总')  # 插入一列数据
            usercount=len(pivoted_df_user)+7
            writer = pd.ExcelWriter(excelName)  # 找到需要写入的excel表格
            sheet_df.to_excel(writer, 'timeLog', index=False)  # 将数据写入到sheet2中   添加几个sheet,需要重新写几次
            sheet_pivoted_df.to_excel(writer, 'Summary',startrow=usercount)  # 将数据写入到sheet1中
            pivoted_df_user.to_excel(writer, 'Sheet1',startrow=3)  # 将数据写入到sheet2中   添加几个sheet,需要重新写几次
            writer.save()  # 保存数据

            oldExcelFile = xlrd.open_workbook(excelName)
            newExcelFile = copy(oldExcelFile)
            timeSheet = oldExcelFile.sheet_by_name('timeLog')
            sumSheet = oldExcelFile.sheet_by_name('Summary')
            oldSheet1 = oldExcelFile.sheet_by_name('Sheet1')

            #获得表格
            newTimeSh = newExcelFile.get_sheet(0)
            newSumSh = newExcelFile.get_sheet(1)
            newSheet1=newExcelFile.get_sheet(2)

            newSumSh.write(3, 1, '')
            newSumSh.write(4, 1, '')
            newSumSh.col(1).width = 256 * 35
            newSheet1.write(3,0,'')  #直接设为空
            newSheet1.write(4,0,'')  #直接设为空
            newSheet1.write(3,1,'')  #直接设为空
            newSheet1.write(4,1,'')  #直接设为空

            # 设置sheet1的总体格式
            for r2 in range(6, oldSheet1.nrows):
                for c2 in range(0, oldSheet1.ncols):
                    cell = oldSheet1.cell_value(r2, c2)
                    newSheet1.write(r2, c2, cell, style3)
            # 设置Summary的总体格式
            for r1 in range(6, sumSheet.nrows):
                for c1 in range(0, sumSheet.ncols):
                    cell = sumSheet.cell_value(r1, c1)
                    newSumSh.write(r1, c1, cell, style3)
            for r3 in range(2,6):
                for c3 in range(0,oldSheet1.ncols):
                    cell=oldSheet1.cell_value(r3,c3)
                    newSheet1.write(r3,c3,cell,style2)
            newSheet1.write(1,0,oldSheet1.cell_value(1,0))

            #将Sheet1中的汇总数据插入的Summary工作簿
            for i in range(3, oldSheet1.nrows):
                for j in range(0, sumSheet.ncols):
                    cell = oldSheet1.cell_value(i, j)
                    newSumSh.write(i, j, cell, style2)

            #设置summary的格式
            for c in range(2, sumSheet.ncols):
                newSumSh.col(c).width = 256 * 8
            for s in range(1, sumSheet.nrows):
                newSumSh.row(s).set_style(tall_style)

            newSumSh.remove_merged_ranges(sumSheet.nrows - 2, sumSheet.nrows - 1, 1, 1)
            newSumSh.write_merge(usercount-1, usercount-1, 0, sumSheet.ncols-1)
            newSumSh.remove_merged_ranges(usercount+1, usercount+1, sumSheet.ncols - 2, sumSheet.ncols - 1)
            newSheet1.remove_merged_ranges(4, 4, oldSheet1.ncols - 2, oldSheet1.ncols - 1)

            newSumSh.write_merge(3, 4, sumSheet.ncols - 1, sumSheet.ncols - 1, '总计',style2)
            newSumSh.write_merge(usercount, usercount+1, sumSheet.ncols - 1, sumSheet.ncols - 1, '总计',style2)
            newSheet1.write_merge(3, 4, oldSheet1.ncols - 1, oldSheet1.ncols - 1, '总计',style2)

            # 设置表头
            name=excelName.split('\\',3)[3]
            depart=name[:-5]
            newSumSh.write_merge(0, 0, 0, oldSheet1.ncols - 1,'')
            newSheet1.write_merge(0, 0, 0, oldSheet1.ncols - 1,'')
            newSumSh.row(0).set_style(fir_tall_style)  # 行高
            newSheet1.row(0).set_style(fir_tall_style)  # 行高
            # newSumSh.row(1).set_style(tall_style)  # 行高

            newSumSh.write(0, 0, depart, style1)
            newSheet1.write(0, 0, depart+'汇总', style1)
            newSumSh.write_merge(1, 1, 0, oldSheet1.ncols - 1, '本月填报缺失人员:',style2)
            newSheet1.write_merge(1, 1, 0, oldSheet1.ncols - 1, '本月填报缺失人员:',style2)
            newSumSh.write_merge(2, 4, 0, 1, '人员/汇总', style2)
            newSumSh.write_merge(usercount, usercount+1, 0, 1, '人员/项目明细', style2)
            # newSheet1.write_merge(2, 4, 0, 1, '')

            # for k in range(usercount, usercount+3):
            #     for z in range(0, sumSheet.ncols):
            #         cell = sumSheet.cell_value(k, z)
            #         newSumSh.write(k, z, cell, style2)

            newSumSh.write_merge(2, 2, 2, oldSheet1.ncols - 1, '活动', style2)
            newSheet1.write_merge(2, 2, 2, oldSheet1.ncols - 1, '活动', style2)

            #设置格式
            for fm in range(0, timeSheet.nrows):
                newTimeSh.row(fm).set_style(tall_style)  # 行高
            newTimeSh.col(0).width = 256 * 35  # 列宽
            newTimeSh.col(1).width = 256 * 16  # 列宽
            newTimeSh.col(4).width = 256 * 30  # 列宽
            for fm in range(0,oldSheet1.nrows):
                newSheet1.row(fm).set_style(tall_style)
            for c in range(2,oldSheet1.ncols):
                newSheet1.col(c).width=256*8

            newSheet1.write(0,2,'')
            newSheet1.write(0,3,'')
            newSheet1.write(1,2,'')
            newSheet1.write(1,3,'')
            xlsName=excelName[:-1]
            xlsNames.append(xlsName)
            newExcelFile.save(xlsName)
            # newExcelFile.save(excelName)
        except:print('without data')
    #填报缺失人员名单
    missingPerson(PATH,xlsNames,newdeparts)


'''for xls in xlsNames:
    # 重新读取数据 获得新的格式
    redata = xlrd.open_workbook(xls)
    newdata = copy(redata)
    sumsheet = redata.sheet_by_index(1)
    sheet = redata.sheet_by_index(2)
    newsheet = newdata.get_sheet(1)
    # sheet1=newdata.get_sheet(2)
    for i in range(3, 5):
        for j in range(2, sumsheet.ncols - 2):
            cell = sumsheet.cell_value(sheet.nrows + i - 2, j)
            newsheet.write(i, j, cell)
    newdata.save(xls)'''



