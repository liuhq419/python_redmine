from redmine import Redmine
from datetime import datetime
from createInsertExcel import insertIntoExcel,createExcel
import os,sys
PATH=os.path.abspath(os.path.join(os.path.dirname(__file__),os.path.pardir))
sys.path.append(PATH)

redmine=Redmine('http://pm.dpark.com.cn/',key='885575e983a0fd543048f2ab10c5d0270f4b1bdd')

#按照不同的部门导出耗时表
def getHourSpend(depart,start_time,end_time,excelName):
    users = redmine.user.all()
    i=0
    for user in users:
        issues = user.issues
        #issues.filter(sort='update_on:desc')
        custom_fields=user.custom_fields
        # print(dir(user))
        for issue in issues:
            # print(dir(issue))
            i_custom_fields = issue.custom_fields  #自定义字段

            project = issue.project  # 项目名称
            update_time=issue.updated_on   #更新日期
            author=issue.author    #用户
            # activity=i_custom_fields[1]    #活动  任务类型 3

            subject = issue.subject   #问题  主题
            #注释  暂时不写
            difficultyLevel=i_custom_fields[0]   #难易程度 1

            time_entries=issue.time_entries  #时间列表
            #print(len(time_entries))
            if custom_fields[0].value[0] == depart:
                # department = custom_fields[0].value[0]  # 部门名称
                if datetime.strptime(start_time, '%Y-%m-%d') <= update_time <= datetime.strptime(end_time, '%Y-%m-%d'):
                    update_time=datetime.strftime(update_time,'%Y/%m/%d')
                # 获得耗时
                    sum=[]
                    for t in time_entries:
                        # print(dir(t))
                        i+=1
                        project1=t.project
                        issue1=t.issue
                        updated_on=t.updated_on
                        user1=t.user
                        hour = t.hours
                        comment=t.comments
                        activity1=t.activity
                        # print(str(project),update_time,str(author),activity.value,str(subject),hour,difficultyLevel.value)
                        #content=[str(project),update_time,str(author),activity.value,str(subject),comment,hour,difficultyLevel.value]
                        content=[str(project1),update_time,updated_on,str(user1),str(activity1),str(issue1),comment,hour,difficultyLevel.value]
                        print(content,depart)
                        if sum==content:
                            continue
                        sum = content
                        #insertIntoExcel(i,content,excelName)   #结果插入到表格
            # fs=issue.custom_fields
            # us=user.custom_fields
            # print(len(fs),len(us))

#获得所有部门耗时，按部门输出
def exportAllDepartTime():
    start_time ='2016-06-01 00:00:00'
    end_time ='2016-06-30 23:59:59'
    department = [ 'GIS平台部','GIS应用部', '行业服务部']#, '房地信息事业部', '交通信息事业部', '信息二部', '信息三部',
                  #'质控部', '规划信息事业部', '暂无部门']  #

    for depart in department:
        excelName=createExcel(depart,PATH)
        getHourSpend(depart,start_time,end_time,excelName)

exportAllDepartTime()

