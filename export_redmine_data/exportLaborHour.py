# -*- coding:utf-8 -*-
from redmine import Redmine
from datetime import datetime
from public_script.excelOpt import *
from config import *
import xlrd,xlwt
from xlutils.copy import copy
import plotly as plt
import numpy as np
import pandas as pd
import os,sys,re

import plotly.graph_objs as go
import plotly.plotly as py

# PATH=os.path.abspath(os.path.join(os.path.dirname(__file__),os.path.pardir))
PATH=os.path.dirname(__file__)
sys.path.append(PATH)

# departments = [ 'GIS平台部','GIS应用部', '房地信息事业部','规划信息事业部', '交通信息事业部',
#                '行业服务部', '暂无部门']
newdeparts=[ 'GIS平台部','GIS应用部',  '不动产信息部','规划信息部',  '交通信息部','咨询服务部']

'''按照不同的部门导出耗时表
获得所有部门耗时，按部门输出,一次性输出所有部门的数据'''

redmine = Redmine(URL, key=MYKEY)
lists = ['项目', '日期', '用户', '活动', '主题', '耗时', '难易度']
allStaffList=['项目', '日期', '用户', '活动', '主题', '耗时', '难易度','所属部门']

def exportLaborHour():
    excelNames = []
    # 首先建立所有部门的数据模板
    excel=createExcel(date,PATH,allStaffList,'所有人员')
    # allStaffExcel=createExcel(date,PATH,allStaffList,'所有人员')   #创建所有人员的耗时表
    for depart in newdeparts:
        excelName = createExcel(date, PATH, lists, depart)
        excelNames.append(excelName)
    excelNames.append(excel)

    #获得所有的redmine系统数据
    time_entries=redmine.time_entry.filter(spent_on=DATE)
    for time_entry in time_entries:
        project=time_entry.project  #项目名称
        spent_on=time_entry.spent_on   #更新日期
        spent_on = datetime.strftime(spent_on,'%Y/%m/%d')  #改变日期格式
        user=time_entry.user   #用户
        if str(user) =='孙虹' or str(user) =='刘曦雯':
            continue
        activity=time_entry.activity  #活动类型
        userid = user.id
        curUser = redmine.user.get(userid)    #用户
        custom_fields = curUser.custom_fields  # user的自定义字段
        department = custom_fields[0].value[0]  # 获得部门名称
        #尝试获取问题的主题，有些没有主题是，所以需要捕获异常
        try:
            # issue=time_entry.issue
            # issueId=issue.id
            iss=redmine.issue.get(time_entry.issue.id)
            curIssue=str(iss)   #获得问题主题字段
            iss_custom_fields=iss.custom_fields
            difficultyLevel=iss_custom_fields[0].value  #获得困难程度字段
        except:
            curIssue=''
            difficultyLevel=''
        hours=time_entry.hours  #耗时
        content = [str(project), spent_on, str(user), str(activity), curIssue, hours, difficultyLevel]

        if department=='房地信息事业部':
            department='不动产信息部'
        elif department=='规划信息事业部':
            department='规划信息部'
        elif department=='交通信息事业部':
            department='交通信息部'
        elif department=='行业服务部':
            department='咨询服务部'
        else:
            department=department
        allContent = [str(project), spent_on, str(user), str(activity), curIssue, hours, difficultyLevel,department]
        insertIntoExcel(allContent,PATH,date,'所有人员')
        if department!='信息三部' and department!='信息二部' and department!='质控部':
            insertIntoExcel(content,PATH,date,department)   #结果插入到表格

    readExcel(excelNames, PATH, newdeparts)

    print('人员耗时情况表导出完毕')
# exportLaborHour()


#输出人员工作饱和度数据和图表
def workTimeLevel(workDays):
    # 获取某个文件下下的所有.xls文件
    fileDir = PATH + '\\exportResult\\PM系统工时填报'
    dirPath = os.listdir(fileDir)  # 获得当前文件夹下的所有文件
    excelNames = []
    for i in range(0, len(dirPath)):
        if re.search('.xls', dirPath[i]):  # 如果有匹配的，则为True
            excelName = dirPath[i]
            excelName = fileDir + '\\' + excelName
            excelNames.append(excelName)  # 获得所有.xls

    for excelName in excelNames:
        oldExcel=xlrd.open_workbook(excelName,formatting_info=True)
        sumSheet=oldExcel.sheet_by_index(1)
        # sheet1=oldExcel.sheet_by_index(2)
        newExcel=copy(oldExcel)
        newSheet1=newExcel.get_sheet(2)

        strMissMan=sumSheet.cell_value(1,0)
        listMissMan=strMissMan.split(', ')
        del listMissMan[0]   #删除下标为0 的数据  #获得缺失填报人员名单
        missWorkHours = []
        for lmm in range(0,len(listMissMan)):
            missWorkHours.append(0.0)    #如果人员没有填写，则工时为0

        huizong=sumSheet.col_values(1)
        wroteManCount=0

        for hz in huizong:  #8/6 10:42
            if hz=='汇总':
                wroteManCount+=1   #得到填写人员的数量
        listWroteMan=sumSheet.col_values(0,6,wroteManCount+5)  #填写人员姓名列表
        wroteWorkHours=sumSheet.col_values(sumSheet.ncols-1,6,wroteManCount+5)  #人员工时列表
        # department=sumSheet.col_values(2,6,wroteManCount+5)
        departMen=listWroteMan+listMissMan   #部门所有人员
        workHours=wroteWorkHours+missWorkHours  #部门所有人员的工时

        newSheet1.write_merge(0, 0, 0, 3, '员工工作饱和度', style1)  # 填写标题
        title = ['姓名', '有效工时', '工作饱和度', '饱和度等级']
        for t in range(0, 4):
            newSheet1.write(1, t, title[t], style2)  # 填写表头
            newSheet1.col(t).width = 256 * 15

        for dm in range(0,len(departMen)):
            newSheet1.write(dm+2,0,departMen[dm],style3)  #姓名
        for wkh in range(0,len(workHours)):
            workHour=workHours[wkh]
            newSheet1.write(wkh+2,1,workHour,style3)   #工时
            baohedu=float(workHour)/(7.5* workDays)  #计算饱和度
            # baohedu=round(baohedu,4)   #取两位小数
            bhd=format(baohedu,'.2%')   #取两位小数,百分比  变为字符串类型
            # print(type(workHour),workHour,baohedu)
            newSheet1.write(wkh+2,2,bhd,style3)   #输入饱和度百分比
            if baohedu >=1:
                bhLevel='饱和'
                style=styBH1
            elif baohedu>=0.9:
                bhLevel='饱和'
                style=styBH2
            elif baohedu>=0.6:
                bhLevel='基本饱和'
                style=styBH2
            else:
                bhLevel='不饱和'
                style=styBH3
            newSheet1.write(wkh+2,3,bhLevel,style)   #输入饱和度等级

            newSheet1.row(0).set_style(fir_tall_style)  # 设置行高
            newSheet1.row(1).set_style(tall_style)  # 设置行高
            newSheet1.row(4).set_style(tall_style)  # 设置行高
            newSheet1.row(wkh+2).set_style(tall_style)  # 设置行高

        newExcel.save(excelName)
    #插入图表
    for excel in excelNames:
        excelData=xlrd.open_workbook(excel,formatting_info=True)
        newData=copy(excelData)
        oldSheet=excelData.sheet_by_index(-1)
        newSheet=newData.get_sheet(-1)
        if '所有人员' in excel:  #8/6 10:40
            continue
        df = pd.ExcelFile(excel)
        df_sheet = df.parse('工作饱和度', skiprows=1)

        trace0 = go.Bar(   #Scatter
            x=df_sheet.姓名,
            y=df_sheet.工作饱和度,
            # mode='lines',
            name='工作饱和度',
            marker=dict(
                color='rgb(158, 202, 225)',
                line=dict(
                    color='rgb(8, 48, 107)',  #color=rgba(184,255,50,0.9)
                    width=1,
                )
            )
        )
        line1=[]
        line2=[]
        line3=[]
        for i in range(1,oldSheet.nrows):
            line1.append(100)
            line2.append(90)
            line3.append(60)

        trace1=go.Scatter(x=df_sheet.姓名, y=line1, mode='lines', name='饱和度=100%',
                              line=dict(color='rgba(0, 255, 0, 1.0)', width=1,),
                          hoverinfo='100%')
        trace2=go.Scatter(x=df_sheet.姓名, y=line2, mode='lines', name='饱和度=90%',
                          line=dict(color='rgba(0, 0, 0, 1.0)', width=1, ),hoverinfo='90%'
                          )
        trace3=go.Scatter(x=df_sheet.姓名, y=line3, mode='lines', name='饱和度=60%',
                          line=dict(color='rgba(255, 0, 0, 1.0)', width=1, ),hoverinfo='60%')
        data = [trace0,trace1,trace2,trace3]
        layout_comp = go.Layout(
            title='员工工作饱和度状况图',
            hovermode='closest',
            xaxis=dict(
                title='姓名',
                ticklen=5,
                # tickfont='Arial',
                tickfont=dict(
                    family='Arial',
                    size=14,
                    color='rgb(0,0,0)',),  #设置字体格式
                zeroline=True,
                gridwidth=2,
                tickcolor='rgb(0,0,0)',
            ),
            yaxis=dict(
                title='饱和度(%)',
                ticks='outside',
                # ticks='inside',
                ticklen=5,  #设置刻度长度
                gridwidth=2,
                zeroline=True,
                tickcolor='rgb(0,0,0)'
            ),
        )
        fig_comp = go.Figure(data=data, layout=layout_comp)
        py.image.save_as(fig_comp, 'workHour.png')  # 将统计图表保存为图片格式
        # py.plot(fig_comp,'my_plot')

        #下一步将图片 插入到excel中
        from PIL import Image
        Image.open('workHour.png').convert("RGB").save('workHour.bmp')
        newSheet.insert_bitmap('workHour.bmp',3,7)
        newData.save(excel)
    os.remove('workHour.png')
    os.remove('workHour.bmp')
    print('工作饱和度图表制作完成')
# workTimeLevel()

#所有员工操作
#员工工作量排序
def sortWorkHours():
    fileDir = PATH + '\\exportResult\\PM系统工时填报\\'
    dirPath = os.listdir(fileDir)  # 获得当前文件夹下的所有文件
    staffSortXls = ''
    for i in range(0, len(dirPath)):
        if re.search('.xls', dirPath[i]):  # 如果有匹配的，则为True
            if '所有人员' in dirPath[i]:
                staffSortXls = fileDir+dirPath[i]
                continue
    # 排序
    try:
        staffSort_df = pd.read_excel(staffSortXls, '工作饱和度', skiprows=[0])
    except:
        staffSort_df = pd.read_excel(staffSortXls, '工作量排序', skiprows=[0])
    timeLog_df = pd.read_excel(staffSortXls, 'timeLog')#, skiprows=[0])
    # summary_df = pd.read_excel(staffSortXls, 'Summary', skiprows=[0])
    # print(staffSort_df)
    writer = pd.ExcelWriter(staffSortXls)
    sorts_df=staffSort_df.sort_values(by=['有效工时'],ascending=False)  # 排序

    timeLog_df.to_excel(writer, 'timeLog', index=False, startrow=0)
    sorts_df.to_excel(writer, '工作饱和度', index=False, startrow=1)
    writer.save()
    allData = xlrd.open_workbook(staffSortXls,formatting_info=True)
    newAllExcel = copy(allData)
    rename = allData.sheet_names().index('工作饱和度')
    newAllExcel.get_sheet(rename).name = '工作量排序'
    timeSheet = allData.sheet_by_index(0)
    sortSheet = allData.sheet_by_index(1)
    newTime = newAllExcel.get_sheet(0)
    newSort = newAllExcel.get_sheet(1)
    # 排序后需要设置格式
    for r1 in range(0, timeSheet.nrows):
        newTime.row(r1).set_style(tall_style)  # 行高
        for c1 in range(0, timeSheet.ncols):
            newTime.write(r1, c1, timeSheet.cell_value(r1, c1), style4)

    for r2 in range(1, sortSheet.nrows):
        newSort.row(r2).set_style(tall_style)  # 行高
        for c2 in range(0, sortSheet.ncols):
            newSort.col(c2).width = 256 * 18  # 列宽
            newSort.write(r2, c2, sortSheet.cell_value(r2, c2), style3)
    newSort.write_merge(0,0,0,3,'所有人员工作量排序',style1)
    newTime.col(0).width = 256 * 38  # 列宽
    newTime.col(1).width = 256 * 16  # 列宽
    newTime.col(4).width = 256 * 30  # 列宽
    newTime.col(7).width = 256 * 16  # 列宽
    newAllExcel.save(staffSortXls)

    df1=sorts_df.head(n=10)  # 查看前几行的数据,默认前5行
    df2=sorts_df.tail(n=10)  # 查看后几行的数据,默认后5行
    df=df1.append(df2)

    # 画图
    trace0 = go.Bar(  # Scatter
        x=df.姓名,
        y=df.有效工时,
        # mode='lines',
        name='工作量',
        marker=dict(
            color='rgb(158, 202, 120)',
            line=dict(
                color='rgb(8, 48, 107)',  # color=rgba(184,255,50,0.9)
                width=1,
            )
        )
    )
    layout_comp = go.Layout(
        title='前十名和最后十名人员工作量详情',
        hovermode='closest',
        xaxis=dict(
            title='姓名',
            ticklen=5,
            # tickfont='Arial',
            tickfont=dict(
                family='Arial',
                size=14,
                color='rgb(0,0,0)', ),  # 设置字体格式
            zeroline=True,
            gridwidth=2,
            tickcolor='rgb(0,0,0)',
        ),
        yaxis=dict(
            title='工作量（小时)',
            ticks='outside',
            # ticks='inside',
            ticklen=5,  # 设置刻度长度
            gridwidth=2,
            zeroline=True,
            tickcolor='rgb(0,0,0)'
        ),
    )
    data=[trace0]
    fig_comp = go.Figure(data=data, layout=layout_comp)
    py.image.save_as(fig_comp, 'workHour.png')  # 将统计图表保存为图片格式

    # 下一步将图片 插入到excel中
    from PIL import Image
    Image.open('workHour.png').convert("RGB").save('workHour.bmp')
    newSort.insert_bitmap('workHour.bmp', 3, 7)
    newAllExcel.save(staffSortXls)

    os.remove('workHour.png')
    os.remove('workHour.bmp')

    print('所有人员工作量排序完成')

if __name__=='__main__':
    exportLaborHour()   #导出耗时数据
    workTimeLevel(workDays=22)   #计算饱和度，并绘制饱和度图表
    sortWorkHours()  #所有的员工工作量排序，并进行绘制图表


