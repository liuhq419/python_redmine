# -*- coding: utf-8 -*-
'''函数功能：创建excel；插入数据到excel'''
import os,sys
import xlrd,xlwt
from xlutils.copy import copy

#设置单元格的样式
# def setExcelStyle():
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
font1.name = 'Arial'
font1.bold = True
font1.height = 0x00140   #字体大小为十六进制转为十进制 除以20   0x00104=320
style1 = xlwt.XFStyle()  # create the style
style1.font = font1
style1.alignment=alignment
style1.borders=borders

#font2为 宋体 加粗 10号
font2 = xlwt.Font()
font2.name = 'Arial'
font2.bold = True
font2.height = 0x00DC  # 字体大小为十六进制转为十进制 除以20   0x00DC=220
style2 = xlwt.XFStyle()  # create the style
style2.font = font2
style2.alignment=alignment
style2.borders=borders

#font3为 宋体  10号
font3 = xlwt.Font()
font3.name = 'Arial'
font3.bold = False
font3.height = 0x00DC  # 字体大小为十六进制转为十进制 除以30   0x00DC=220
style3 = xlwt.XFStyle()  # create the style
style3.font = font3
style3.alignment=alignment  #设置对齐方式
style3.borders=borders


style4 = xlwt.XFStyle()  # create the style
style4.font = font3
style4.alignment=alignment1  #设置对齐方式
style4.borders=borders

    # return (style1,style2,style3,style4)
#三种不同的格式，公共变量
# style1, style2, style3,style4 = setExcelStyle()
#行高
tall_style=xlwt.easyxf('font:height 280;')  #普通行行高
fir_tall_style=xlwt.easyxf('font:height 400;')  #普通行行高

#设置字体的函数，  corlorId=3：绿色 ；2：红色 1:白色  0：黑色
def setFont(colorId):
    fontBH = xlwt.Font()
    fontBH.name = 'Arial'
    fontBH.bold = False
    fontBH.height = 0x00DC  # 字体大小为十六进制转为十进制 除以30   0x00DC=220
    fontBH.colour_index=colorId
    return fontBH

#设置边框函数
def setBorders(leftid,rightid,topid,bottomid):
    # 设置边框
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    return borders

#设置工作饱和度写入表格的格式
styBH1 = xlwt.XFStyle()   #过于饱和 大于100%
styBH1.font=setFont(3)     #宋体10号   #绿色
styBH1.borders=setBorders(1,1,1,1)   #有边框
styBH1.alignment=alignment   #居中对齐

styBH2=xlwt.XFStyle()  #饱和与基本饱和 大于90% 和介于60%-90% 都是为黑
styBH2.font=setFont(0)
styBH2.borders=setBorders(1,1,1,1)
styBH2.alignment=alignment

styBH3=xlwt.XFStyle()   #饱和度小于60%
styBH3.font=setFont(2)   #红色
styBH3.borders=setBorders(1,1,1,1)
styBH3.alignment=alignment



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
        timeSheet.write(0,i,lists[i],style2)    # 逐个插入lists列表中的数据，即为表头

        # if str=='PM系统工时填报':
        #     timeSheet.col(i).width=256*14
        #     timeSheet.col(0).width=256*32
        #     timeSheet.col(4).width=256*32
        # else:
        #     timeSheet.col(i).width=256*10
        #     timeSheet.col(1).width=256*32
        #     timeSheet.col(4).width=256*32
        #     timeSheet.col(6).width=256*20
        #     timeSheet.col(7).width=256*20
        #     if depart:
        #         timeSheet.col(9).width=256*20
        #     else:
        #         timeSheet.col(8).width=256*20
        #         timeSheet.col(10).width=256*20
        #
        # timeSheet.row(0).set_style(tall_style)

    filePath = PATH + '\\exportResult\\'
    excelName= str+'\\'+depart + str + date + '.xls'
    pathExcelName = filePath +excelName
    file.save(pathExcelName)  # 保存excel
    return pathExcelName

'''插入数据到excel中
context 为所要插入的数据，'''
def insertIntoExcel(content,PATH,date,depart='',str= 'PM系统工时填报'):
    filePath = PATH + '\\exportResult\\'
    excelName = filePath +str+'\\'+depart +str+ date + '.xls'
    #打开要插入的表，并将数据复制到新的表中
    oldExcel=xlrd.open_workbook(excelName, formatting_info=True)
    newExcel=copy(oldExcel)
    sheet=oldExcel.sheet_by_index(0)
    newSheet=newExcel.get_sheet(0)

    #逐行插入数据
    r=sheet.nrows
    for col in range(0,sheet.ncols):
        #table.nrows控制始终将数据插入到当前文件的最后一行
        # newSheet.row(r-1).set_style(tall_style)
        # newSheet.row(r).set_style(tall_style)
        newSheet.write(r,col,content[col],style4)
    newExcel.save(excelName)


def missingPerson(PATH,excelNames,newdeparts,staffExcel='格网各部门人员清单.xls'):
    staffxls=PATH+'\\public_script\\'+staffExcel
    newdeparts.append('所有人员')
    for depart in newdeparts:
        staffXls=xlrd.open_workbook(staffxls)
        staffSht=staffXls.sheet_by_name(depart)
        totalStaff=staffSht.col_values(0,start_rowx=1)
        missingStaff=[]
        for xlsName in excelNames:
            if depart in xlsName:
                timeXls=xlrd.open_workbook(xlsName,formatting_info=True)
                newTimeXLs=copy(timeXls)
                timeSht=timeXls.sheet_by_index(1)
                sheet1=timeXls.sheet_by_index(2)
                newTimeSht1=newTimeXLs.get_sheet(1)
                newsheet1=newTimeXLs.get_sheet(2)
                wroteStaff=timeSht.col_values(0,0,sheet1.nrows+1)
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
                        # newTimeSht.write(1,0,miss)
                        newTimeSht1.write(1,0,miss,style4)
                for i in range(0,sheet1.nrows):
                    for j in range(0,sheet1.ncols):
                        newsheet1.write(i,j,'')
                try:
                    newsheet1.remove_merged_ranges(1, 1, 2,4)
                except:pass
                try:
                    newsheet1.remove_merged_ranges(1, 1, 3,4)
                except:pass
                try:
                    newsheet1.remove_merged_ranges(1,1,2,5)
                    newsheet1.remove_merged_ranges(1,1,3,5)
                except:pass
                newTimeXLs.save(xlsName)
            else:continue
# missingPerson()

import pandas as pd
import numpy as np

# 读取excel表格中的数据，并进行相应的处理，本例是添加透视表
def readExcel(excelNames,PATH,newdeparts):
    # xlsNames = []
    for excelName in excelNames:
        excel_df = pd.ExcelFile(excelName)#,engine='xlsxwriter')  # 获得excel
        sheet_df = excel_df.parse('timeLog')  # 获得excel的工作簿   #返回dataFrame
        sheet_df = sheet_df.fillna('-')  # 空值处理，将空值用''填充
        try:
            # 对数据表进行添加透视表
            sheet_pivoted_df = pd.pivot_table(sheet_df, values='耗时', index=['用户', '项目'], columns=['活动', '难易度'],
                                              aggfunc={'耗时': np.sum}, fill_value='', margins=True,
                                              margins_name='总计')
            # pivoted_df_user = pd.pivot_table(sheet_df, values='耗时', index=['用户'], columns=['活动', '难易度'],
            #                                  aggfunc={'耗时': np.sum}, fill_value='', margins=True,
            #                                  margins_name='总计')
            # pivoted_df_user.insert(0, '活动', '汇总')  # 插入一列数据  8/6 10:40
            if '所有人员' not in excelName:
                pivoted_df_user = pd.pivot_table(sheet_df, values='耗时', index=['用户'], columns=['活动', '难易度'],
                                                 aggfunc={'耗时': np.sum}, fill_value='', margins=True,
                                                 margins_name='总计')
                pivoted_df_user.insert(0, '活动', '汇总')  # 插入一列数据
            else:
                pivoted_df_user = pd.pivot_table(sheet_df, values='耗时', index=['用户','所属部门'], columns=['活动', '难易度'],
                                                 aggfunc={'耗时': np.sum}, fill_value='', margins=True,
                                                 margins_name='总计')
            usercount=len(pivoted_df_user)+7
            writer = pd.ExcelWriter(excelName)  # 找到需要写入的excel表格
            sheet_df.to_excel(writer, 'timeLog', index=False)  # 将数据写入到sheet2中   添加几个sheet,需要重新写几次
            sheet_pivoted_df.to_excel(writer, 'Summary',startrow=usercount)  # 将数据写入到sheet1中
            pivoted_df_user.to_excel(writer, 'Summary',startrow=3)  # 将数据写入到sheet2中   添加几个sheet,需要重新写几次
            pivoted_df_user.to_excel(writer, '工作饱和度',startrow=1)  # 将数据写入到sheet2中   添加几个sheet,需要重新写几次
            writer.save()  # 保存数据
            #读取表格
            oldExcelFile = xlrd.open_workbook(excelName,formatting_info=True)
            newExcelFile = copy(oldExcelFile)
            timeSheet = oldExcelFile.sheet_by_name('timeLog')
            sumSheet = oldExcelFile.sheet_by_name('Summary')
            #获得复制后的表格工作簿
            newTimeSh = newExcelFile.get_sheet(0)
            newSumSh = newExcelFile.get_sheet(1)

            #设置格式timeLog的格式
            for fm in range(0, timeSheet.nrows):
                newTimeSh.row(fm).set_style(tall_style)  # 行高
            newTimeSh.col(0).width = 256 * 38  # 列宽
            newTimeSh.col(1).width = 256 * 16  # 列宽
            newTimeSh.col(4).width = 256 * 30  # 列宽

            #设置summary的格式
            for c in range(2, sumSheet.ncols):
                newSumSh.col(c).width = 256 * 8  #列宽
            newSumSh.col(1).width = 256 * 40  #第二列列宽
            newSumSh.row(0).set_style(fir_tall_style)  # 首行行高
            for s in range(1, sumSheet.nrows):
                newSumSh.row(s).set_style(tall_style)   #行高

            newSumSh.remove_merged_ranges(sumSheet.nrows - 2, sumSheet.nrows - 1, 1, 1)
            newSumSh.remove_merged_ranges(4, 4, sumSheet.ncols - 2, sumSheet.ncols - 1)
            newSumSh.remove_merged_ranges(usercount+1, usercount+1, sumSheet.ncols - 2, sumSheet.ncols - 1)
            if '所有人员' in excelName:
                try:
                    newSumSh.remove_merged_ranges(usercount-3, usercount-2, 1,1)
                except:pass

            newSumSh.write_merge(3, 4, sumSheet.ncols - 1, sumSheet.ncols - 1, '总计',style2)
            newSumSh.write_merge(usercount, usercount+1, sumSheet.ncols - 1, sumSheet.ncols - 1, '总计',style2)

            # 设置表头
            name=excelName.split('\\',3)[3]
            depart=name[:-4]
            newSumSh.write_merge(0, 0, 0, sumSheet.ncols - 1,'')

            newSumSh.write(0, 0, depart, style1)
            newSumSh.write_merge(1, 1, 0, sumSheet.ncols - 1,'')
            newSumSh.write_merge(2, 4, 0, 1, '人员/汇总', style2)
            newSumSh.write_merge(usercount, usercount+1, 0, 1, '人员/项目明细', style2)

            newSumSh.write_merge(2, 2, 2, sumSheet.ncols - 1, '活动', style2)

            newExcelFile.save(excelName)

        except:
            print(excelName+' 没有数据')
        #对所有人员表进行处理，获得耗时排序

    #填报缺失人员名单
    # missingPerson(PATH, excelNames, newdeparts)
    try:
        missingPerson(PATH,excelNames,newdeparts)
    except:
        pass

#暂时无用
def allStaffStatistic(allStaffExcel):#excelNames):
    # allStaffExcel=''
    # for name in excelNames:
    #     if '所有人员PM系统工时填报201607.xls' in name:
    #         allStaffExcel=name
    #         continue
    excel_df = pd.ExcelFile(allStaffExcel)  # ,engine='xlsxwriter')  # 获得excel
    sheet_df = excel_df.parse('timeLog')  # 获得excel的工作簿   #返回dataFrame
    sheet_df = sheet_df.fillna('-')  # 空值处理，将空值用''填充
    sheet_pivoted_df = pd.pivot_table(sheet_df, values='耗时', index=['用户','所属部门'],columns='难易度',
                                     aggfunc={'耗时': np.sum}, fill_value='', margins=True,
                                     margins_name='总计')
    sheet_pivoted_df=sheet_pivoted_df.sort_values(by='耗时')   #耗时排序
    writer = pd.ExcelWriter(allStaffExcel)  # 找到需要写入的excel表格
    sheet_df.to_excel(writer, 'timeLog', index=False)  # 将数据写入到sheet中   添加几个sheet,需要重新写几次
    sheet_pivoted_df.to_excel(writer, '人员耗时排序', startrow=2)  # 将数据写入到sheet中
    writer.save()  # 保存数据

    oldExcelFile = xlrd.open_workbook(allStaffExcel, formatting_info=True)
    newExcelFile = copy(oldExcelFile)
    timeSheet = oldExcelFile.sheet_by_name('timeLog')
    timeSort = oldExcelFile.sheet_by_name('人员耗时排序')
    # 获得复制后的表格工作簿
    newTimeSh = newExcelFile.get_sheet(0)
    newTSort = newExcelFile.get_sheet(1)

    # 设置格式newTSort的格式
    for i in range(1,timeSort.nrows):
        newTSort.row(i).set_styel(tall_style)     #行高
    newTSort.row(0).set_styel(fir_tall_style)  # 行高
    newTSort.write_merge(0, 0, 0,3, '人员工作量排序', style2)

    # 设置格式timeLog的格式
    for fm in range(0, timeSheet.nrows):
        newTimeSh.row(fm).set_style(tall_style)  # 行高
    newTimeSh.col(0).width = 256 * 38  # 列宽
    newTimeSh.col(1).width = 256 * 16  # 列宽
    newTimeSh.col(4).width = 256 * 30  # 列宽
    newTimeSh.col(7).width = 256 * 16  # 列宽

import re
def getExcelNames(PATH):
    # 获取某个文件下下的所有.xls文件
    fileDir = PATH + '\\exportResult\\PM系统工时填报'
    dirPath = os.listdir(fileDir)  # 获得当前文件夹下的所有文件
    excelNames = []
    for i in range(0, len(dirPath)):
        if re.search('.xls', dirPath[i]):  # 如果有匹配的，则为True
            excelName = dirPath[i]
            excelName = fileDir + '\\' + excelName
            excelNames.append(excelName)  # 获得所有.xls



