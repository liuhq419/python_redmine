from redmine import Redmine
from datetime import datetime
# from createInsertExcel import insertIntoExcel,createExcel
from createInsertExcel import *
import xlrd,xlwt
from xlutils.copy import copy
import os,sys
PATH=os.path.abspath(os.path.join(os.path.dirname(__file__),os.path.pardir))
sys.path.append(PATH)

redmine=Redmine('http://pm.dpark.com.cn/',key='885575e983a0fd543048f2ab10c5d0270f4b1bdd')


#按照不同的部门导出耗时表
#获得所有部门耗时，按部门输出,一次性输出所有部门的数据
def getHourSpend():
    '''按照spent_on进行过滤，导出什么月份的数据需要进行相应的修改'''
    filterDate='><2016-06-01|2016-06-30'   #过滤日期,即为所筛选的日期范围
    date='201606'  #excel表格名称的一部分，导出的是哪一个月需要进行修改

    departments = [ 'GIS平台部','GIS应用部', '行业服务部', '房地信息事业部', '交通信息事业部',
                   '信息二部', '信息三部','质控部', '规划信息事业部', '暂无部门']
    excelNames=[]
    pathExcelNames=[]
    #首先建立所有部门的数据模板
    for depart in departments:
        excelName,pathExcelName=createExcel(depart,date,PATH)
        excelNames.append(excelName)   #生成的excel表格的名称
        pathExcelNames.append(pathExcelName)   #带有路径的excel表格的名称

    def getAndInsertData(sortType='',sheetIndex=0):
        #获得所有的redmine系统数据
        time_entries=redmine.time_entry.filter(spent_on=filterDate,sort=sortType,limit=100)
        for time_entry in time_entries:
            project=time_entry.project  #项目名称
            spent_on=time_entry.spent_on   #更新日期
            spent_on = datetime.strftime(spent_on,'%Y/%m/%d')  #改变日期格式
            user=time_entry.user   #用户
            activity=time_entry.activity  #活动类型
            userid = user.id
            curUser = redmine.user.get(userid)    #用户
            custom_fields = curUser.custom_fields  # user的自定义字段
            department = custom_fields[0].value[0]  # 获得部门名称
            #尝试获取问题的主题，有些没有主题是，所以需要捕获异常
            try:
                iss=redmine.issue.get(time_entry.issue.id)
                curIssue=str(iss)   #获得问题主题字段
                iss_custom_fields=iss.custom_fields
                difficultyLevel=iss_custom_fields[0].value  #获得困难程度字段
            except:
                curIssue=''
                difficultyLevel=''
            hours=time_entry.hours  #耗时
            content = [str(project), spent_on, str(user), str(activity), curIssue, hours, difficultyLevel]
            for depart in departments:    #分部门插入到表格中
                if department==depart:
                    insertIntoExcel(content,department,PATH,date,sheetIndex)   #结果插入到表格
    #将工时数据插入到表格，默认按照时间进行排序
    getAndInsertData()
    #将工时数据插入到表格，按照用户名进行排序升序排序
    getAndInsertData(sortType='user:asc',sheetIndex=1)

    #将已排序的人员工时情况表进行处理，插入到项目耗时情况表中
    def operateLaborHour(excelNames,pathExcelNames):
        excelid = -1
        #逐表格操作
        for pathExcelName in pathExcelNames:
            excelid+=1
            name=excelNames[excelid][0:-18]  #取部门名称
            excel=xlrd.open_workbook(pathExcelName, formatting_info=True)
            sortSheet=excel.sheet_by_index(1)   #得到sortSheet工作簿，人员耗时表
            newExcel = copy(excel)
            proSheetTemp=newExcel.add_sheet('sheet1',cell_overwrite_ok=True)  #添加项目耗时工作簿
            activityTypes=[]
            diffLevels=['简单','普通','复杂']
            activeCol=sortSheet.col_values(colx=3,start_rowx=1)  #读取sortSheet中的一列数据
            for activityType in activeCol:  #遍历活动列
                #判断共有多少个活动类型，并创建相应的表结构
                if activityType not in activityTypes:
                    activityTypes.append(activityType)
                    activityTypes=sorted(activityTypes)  #将重复的过滤
            # 设置sheet1的表头
            setProSheetStyle(activityTypes, proSheetTemp, name, date)
            #读取人员耗时表中的数据
            user=''
            count=4   #设置count行数，从第4行开始插入数据，前几行留给表头
            for r in range(1, sortSheet.nrows):
                proSheetTemp.row(count).set_style(tall_style)  # 设置行高
                perRowValue=sortSheet.row_values(rowx=r)  #读取sortSheet中的一行数据
                project = perRowValue[0]
                userName = perRowValue[2]
                activityType = perRowValue[3]
                laborHour = perRowValue[5]
                diffLevel = perRowValue[6]
                #根据用户，获得该用户下的所有项目等信息
                if userName!=user:
                    count+=1
                    proSheetTemp.write(count, 0, userName, style3)
                    user = userName
                else:
                    count+=1
                    proSheetTemp.write(count,0,project,style4)
                    ai=activityTypes.index(activityType)
                    di=diffLevels.index(diffLevel)
                    proSheetTemp.write(count,ai*4+di,laborHour,style3)

            # 删除排序好的数据表
            new=newExcel.get_sheet(1)
            for r in range(0, sortSheet.nrows):
                for c in range(0, sortSheet.ncols):
                    new.write(r, c, '')
            #保存数据
            newExcel.save(pathExcelName)  #保存的数据有重复，需要删除重复，然后将工时相加


    operateLaborHour(excelNames,pathExcelNames)

getHourSpend()
