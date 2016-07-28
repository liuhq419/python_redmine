# -*- coding:utf-8 -*-
from redmine import Redmine
from datetime import datetime
from excelOpt import *
import os,sys
PATH=os.path.abspath(os.path.join(os.path.dirname(__file__),os.path.pardir))
sys.path.append(PATH)

redmine = Redmine('http://pm.dpark.com.cn/', key='885575e983a0fd543048f2ab10c5d0270f4b1bdd')
#按照更新时间进行过滤，导出什么月份的数据需要进行相应的修改
filterDate='><2016-06-01|2016-06-30'   #过滤日期,即为所筛选的日期范围
date='201606'  #excel表格名称的一部分，导出的是哪一个月需要进行修改

departments = [ 'GIS平台部','GIS应用部', '行业服务部', '房地信息事业部', '交通信息事业部',
               '信息二部', '信息三部','质控部', '规划信息事业部', '暂无部门']


'''PM系统任务完成情况导出'''
def taskFinishedLevel():
    lists1=['#主题编号','项目','状态','优先级','主题','姓名','更新时间','计划完成时间','% 完成','创建时间']
    lists2=['#主题编号','项目','状态','优先级','主题','姓名','部门','更新时间','计划完成时间','% 完成','创建时间']
    str='PM系统任务完成情况'
    # 首先建立所有部门的数据模板
    for depart in departments:
        createExcel(date,PATH,lists1,depart,str)
    #建立总表
    createExcel(date,PATH,lists2,depart='',str=str)
    #获得所有问题issue
    issues=redmine.issue.filter(updated_on=filterDate,status_id='*')
    for issue in issues:
        issueId=issue.id   #主题编号
        project=issue.project.name   #项目名称
        status=issue.status.name  #主题状态
        priority=issue.priority.name   #优先级
        subject=issue.subject  #主题名称
        author=issue.author   #姓名
        user=redmine.user.get(author.id)  #获得用户
        custom_fields=user.custom_fields  #用户自定义字段
        depart=custom_fields[0].value[0]   #部门名称
        updated_on=issue.updated_on  #更新时间
        due_date=issue.due_date           #计划完成时间
        done_ratio=issue.done_ratio  #%完成
        created_on=issue.created_on  #创建时间
        updated_on = datetime.strftime(updated_on,'%Y/%m/%d %H:%M')  #改变日期格式
        due_date = datetime.strftime(due_date,'%Y/%m/%d %H:%M')  #改变日期格式
        created_on = datetime.strftime(created_on,'%Y/%m/%d %H:%M')  #改变日期格式
        content1=[issueId,project,status,priority,subject,author.name,updated_on,due_date,done_ratio,created_on]
        content2=[issueId,project,status,priority,subject,author.name,depart,updated_on,due_date,done_ratio,created_on]
        try:
            insertIntoExcel(content1,PATH,date,depart,str)
            insertIntoExcel(content2,PATH,date,depart='',str=str)
        except:pass

taskFinishedLevel()
