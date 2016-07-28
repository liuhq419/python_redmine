#coding:utf-8
from redmine import Redmine
from createInsertExcel import insertIntoExcel
import time
from datetime import datetime

redmine=Redmine('http://pm.dpark.com.cn/',key='885575e983a0fd543048f2ab10c5d0270f4b1bdd')
# # redmine=Redmine('http://pm.dpark.com.cn/',username='SDKUser',password='qwert!@#$%')
# print(redmine)

# project=redmine.project.get('sz_onemap')
# #project=redmine.project.all()
# # print(project.id)
# #l=list(redmine.project.all()[0])
# #print(len(l))
# print(project)
# print(list(redmine.project.all()[0]))
# issue=project.issues[0]
# issue_id=issue.id
# print(issue)
# print(issue_id)
#
# # error()使用账户和密码会报相关的错误，可以只用key进行认证
# # redmine.exceptions.AuthError: Invalid authentication details

#获得所有的issue（问题）
# issues=redmine.issue.all()
# print('len_issue:',len(issues))
#
# filter_issues=redmine.issue.filter(status_id='*')
# print(len(filter_issues))
# for i in filter_issues:
#     print(i)

#根据删选条件获得issues
# issues=redmine.issue.filter(entry_time=)
# print(issues)
# groups=redmine.group.all()
# print(groups)
# for group in groups:
#     print(group)
# users=redmine.user.all()
# print(users)
# for user in users:
#     print(user)
#获得自定义字段
# custom_fields=redmine.custom_field.all()[15]
# print(len(custom_fields))
# print(custom_fields)
# for field in custom_fields:
#     print(field)
#获取当前用户
# user=redmine.user.get('current')
# print(user)

#时间过滤
#time=redmine.time_entry.filter(from_date='2016-06-01',to_date='2016-06-30',limit=10)
# for t in time:
#     print(t)

#对issue进行时间和字段的筛选
# issues=redmine.issue.filter(cf_x=0,updated_on='><2016-06-01|2016-06-30')
# print(len(issues))
# for issue in issues:
#     # project_name=redmine.issue.filter()
#     print(issue)
'''users=redmine.user.all()
list=[]
for user in users:
    issues=user.issues
    custom_fields=user.custom_fields
    for issue in issues:
        issue_id=issue.id
        issue_user=user'''

projects=redmine.project.all(limit=5)
for pro in projects:
    print(pro.id)
    project=redmine.project.get(pro.id)
    print(project)
    i=project.issues
    print(i)

    # cates=project.issue_categories
    # print(issues,cates)

    # cates=list(cates)
    # print(type(cates))
    # #print(i)
    print(len(i))
    for issue in i:
        subject=issue.subject
        entry_time=issue.time_entry
        # member=issue.user???
        for time in entry_time:
            hours=time.hours
            update_on=time.updated_on
            print(pro.name,subject,hours,update_on)



users = redmine.user.filter(limit=5,cf_0='GIS平台部')
for user in users:
    # u_custom_fields=user.custom_fields
    # u_custom_category=user.custom_field_value
    # cate=u_custom_category.categories
    # print(cate)
    issues = user.issues

    for issue in issues:
        subject = issue.subject
        time = issue.time_entries
        i_custom_fields = issue.custom_fields

        # 获得耗时
        for t in time:
            hour = t.hours
            update_time = t.updated_on
            if datetime.strptime('2016-06-01','%Y-%m-%d') <= update_time <=datetime.strptime('2016-06-30','%Y-%m-%d'):
                update=update_time
                print(user,subject,hour, update,i_custom_fields[0])




   # list=['a','b','c']
    #insertIntoExcel(str(project),list)  #未知错误，插入数据有误
