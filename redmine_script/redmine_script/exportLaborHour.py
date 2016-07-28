# -*- coding:utf-8 -*-

from datetime import datetime
from excelOpt import *
from setParam import *
import os,sys

# PATH=os.path.abspath(os.path.join(os.path.dirname(__file__),os.path.pardir))
PATH=os.path.dirname(__file__)
sys.path.append(PATH)


departments = [ 'GIS平台部','GIS应用部', '行业服务部', '房地信息事业部', '交通信息事业部',
               '信息二部', '信息三部','质控部', '规划信息事业部', '暂无部门']

'''按照不同的部门导出耗时表
获得所有部门耗时，按部门输出,一次性输出所有部门的数据'''
def exportLaborHour():
    lists = ['项目', '日期', '用户', '活动', '主题', '耗时', '难易度']
    # 首先建立所有部门的数据模板
    for depart in departments:
        createExcel(date,PATH,lists,depart)
    #获得所有的redmine系统数据
    time_entries=redmine.time_entry.filter(spent_on=filterDate)
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
        insertIntoExcel(content,PATH,date,department)   #结果插入到表格

exportLaborHour()

