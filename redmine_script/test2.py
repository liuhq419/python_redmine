from redmine import Redmine
from datetime import datetime
from createInsertExcel import insertIntoExcel,createExcel,readExcel
import xlrd,xlwt
import os,sys
PATH=os.path.abspath(os.path.join(os.path.dirname(__file__),os.path.pardir))
sys.path.append(PATH)

redmine=Redmine('http://pm.dpark.com.cn/',key='885575e983a0fd543048f2ab10c5d0270f4b1bdd')

#按照不同的部门导出耗时表
#获得所有部门耗时，按部门输出,一次性输出所有部门的数据

def getHourSpend():
    startTime='2016/06/01'
    endTime='2016/06/30'
    date='201606'  #excel表格名称的一部分，导出的是哪一个月需要进行修改
    departments = [ 'GIS平台部','GIS应用部', '行业服务部', '房地信息事业部', '交通信息事业部',
                   '信息二部', '信息三部','质控部', '规划信息事业部', '暂无部门']
    excelNames=[]
    #首先建立所有部门的数据模板
    for depart in departments:
        excelName=createExcel(depart,date,PATH)
        excelNames.append(excelName)


    time_entries=redmine.time_entry.all(sort='updated_on:desc')
    for time_entry in time_entries:
        #print(dir(time_entry))
        project=time_entry.project
        updated=time_entry.updated_on
        updated_on = datetime.strftime(updated, '%Y/%m/%d')  #改变日期格式
        user=time_entry.user
        #print(dir(user))
        activity=time_entry.activity
        try:
            iss=redmine.issue.get(time_entry.issue.id)
            curIssue=str(iss)
            iss_custom_fields=iss.custom_fields
            difficultyLevel=iss_custom_fields[0].value
        except:
            curIssue=''
            difficultyLevel=''
        try:
            userid = user.id
            curUser = redmine.user.get(userid)
            custom_fields = curUser.custom_fields  # user的自定义字段
            department = custom_fields[0].value[0]  # 获得部门名称
        except:department='暂无部门'
        hours=time_entry.hours

        if datetime.strptime(startTime,'%Y/%m/%d') <=datetime.strptime(updated_on,'%Y/%m/%d')<=datetime.strptime(endTime,'%Y/%m/%d'):
            content = [str(project), updated_on, str(user), str(activity), curIssue, hours, difficultyLevel]
            print(content, department)
            for depart in departments:
                if department==depart:
                    insertIntoExcel(content,department,PATH,date)   #结果插入到表格
getHourSpend()
