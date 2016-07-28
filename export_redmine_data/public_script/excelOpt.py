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

    #font2为 宋体 加粗 11号
    font2 = xlwt.Font()
    font2.name = '宋体'
    font2.bold = True
    font2.height = 0x00DC  # 字体大小为十六进制转为十进制 除以20   0x00DC=220
    style2 = xlwt.XFStyle()  # create the style
    style2.font = font2
    style2.alignment=alignment
    style2.borders=borders

    #font3为 宋体  11号
    font3 = xlwt.Font()
    font3.name = '宋体'
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

    return (style1,style2,style3,style4)
#三种不同的格式，公共变量
style1, style2, style3,style4 = setExcelStyle()
#行高
tall_style=xlwt.easyxf('font:height 280;')  #普通行行高


'''write_merge(x, x + m, y, y + n, string, sytle)
x表示行，y表示列，m表示跨行个数，n表示跨列个数，string表示要写入的单元格内容，
style表示单元格样式。其中，x，y，m，n，都是以0开始计算的。'''

'''建立excel，存储查询结果'''
'''该函数在insertIntoExcel()插入数据函数中调用，name,filePath,lists在insertExcel()中获得'''
def createExcel(date,PATH,lists,depart='',str= 'PM系统工时填报'):
    # 新建xls，新建名为sheet1的工作簿
    file=xlwt.Workbook()  #encoding='ascii'
    timeSheet=file.add_sheet('Sheet1',cell_overwrite_ok=True)  #添加人员维度的工作簿

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
        newSheet.row(r-1).set_style(tall_style)
        newSheet.row(r).set_style(tall_style)
        newSheet.write(r,col,content[col],style4)
    newExcel.save(excelName)