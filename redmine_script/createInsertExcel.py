# coding: utf-8
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
first_tall=xlwt.easyxf('font:height 540;')  #首行大标题


'''write_merge(x, x + m, y, y + n, string, sytle)
x表示行，y表示列，m表示跨行个数，n表示跨列个数，string表示要写入的单元格内容，
style表示单元格样式。其中，x，y，m，n，都是以0开始计算的。'''

'''建立excel，存储查询结果'''
'''该函数在insertIntoExcel()插入数据函数中调用，name,filePath,lists在insertExcel()中获得'''
def createExcel(name,date,PATH):
    # 新建xls，新建名为sheet1的工作簿
    file=xlwt.Workbook(encoding='ascii')
    #proSheet=file.add_sheet('sheet1',cell_overwrite_ok=True)  #添加项目维度的工作簿
    timeSheet=file.add_sheet('timelog',cell_overwrite_ok=True)  #添加人员维度的工作簿
    sortSheet=file.add_sheet('sheet2',cell_overwrite_ok=True)  #添加已经进行排序的数据
    lists=['项目','日期','用户','活动','主题','耗时','难易度']

    # 设置timelog工作簿的表头
    # timeSheet.set_col_default_width(256 * 16)  # 设置列宽
    timeSheet.row(0).set_style(tall_style)  #设置行高
    sortSheet.row(0).set_style(tall_style)
    for i in range(0,len(lists)):
        timeSheet.col(i).width=256*16
        timeSheet.row(0).set_style(tall_style)
        timeSheet.write(0,i,lists[i],style2)    # 逐个插入lists列表中的数据，即为表头
        sortSheet.col(i).width=256*16
        sortSheet.write(0,i,lists[i],style2)    # 逐个插入lists列表中的数据，即为表头

    filePath = PATH + '\\redmine_script\\exportResult\\'
    excelName= name + 'PM系统工时填报' + date + '.xls'
    pathExcelName = filePath +excelName
    file.save(pathExcelName)  # 保存excel
    return (excelName,pathExcelName)

def setProSheetStyle(actions,proSheet,name,date):
     #设置工作簿proSheet 的格式
    for r in range(1,5):
        proSheet.row(r).set_style(tall_style)  #设置行高
    proSheet.row(0).set_style(first_tall)  #设置行高
    title=name+date[-1]+'月PM数据填报情况'
    for c in range(1,34):
        # proSheet.set_col_default_width(256*2)
        proSheet.col(c).width=256*5
    proSheet.col(0).width=256*36

    colLen = len(actions)*4+1
    proSheet.write_merge(0,0,0,colLen,title,style1)  #第一行合并，写入数据
    proSheet.write_merge(1, 1, 0, colLen, '本月填报缺失人员',style4)  # 第二行合并，写入数据
    proSheet.write_merge(2,4,0,0,'人员/项目',style2)
    proSheet.write_merge(2,2,1,colLen,'活动',style2)
    j=-1
    for i in range(1,colLen-3,4):
        j+=1
        proSheet.write_merge(3,3,i,i+3,actions[j],style2)
    proSheet.write_merge(3,4,colLen,colLen,'总计',style2)
    difficultyLevel=['简单','普通','复杂','汇总']
    k=-1
    for i in range(1,colLen):
        k+=1
        if k==4:
            k=0
        proSheet.write(4,i,difficultyLevel[k],style2)



'''插入数据到excel中
context 为所要插入的数据，flag标签用于指向特定的代码块'''
def insertIntoExcel(content,name,PATH,date,sheetIndex=1):
    filePath = PATH + '\\redmine_script\\exportResult\\'
    excelName = filePath + name + 'PM系统工时填报' + date + '.xls'
    #打开要插入的表，并将数据复制到新的表中
    oldExcel=xlrd.open_workbook(excelName, formatting_info=True)
    newExcel=copy(oldExcel)
    sheet=oldExcel.sheet_by_index(sheetIndex)
    newSheet=newExcel.get_sheet(sheetIndex)

    # newSheet.set_col_default_width(256 * 16)  # 设置列宽
    #逐行插入数据
    r=sheet.nrows
    for col in range(0,sheet.ncols):
        #table.nrows控制始终将数据插入到当前文件的最后一行
        newSheet.col(col).width=256*16
        newSheet.row(r-1).set_style(tall_style)
        newSheet.write(r,col,content[col],style4)
    newExcel.save(excelName)


#弃用
def readExcel(path,name):
    data = xlrd.open_workbook(path+'\\redmine_script\\'+name)
    table = data.sheet_by_index(0)
    timeScope= table.row_values(1)[1]  #获得表中的需要查询的日期范围
    dateValue=table.row_values(1)[0]

    return timeScope,dateValue