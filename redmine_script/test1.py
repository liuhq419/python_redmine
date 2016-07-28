from redmine import Redmine
from datetime import datetime
from createInsertExcel import insertIntoExcel,createExcel,readExcel
import xlrd,xlwt
import os,sys
PATH=os.path.abspath(os.path.join(os.path.dirname(__file__),os.path.pardir))
sys.path.append(PATH)

redmine=Redmine('http://pm.dpark.com.cn/',key='885575e983a0fd543048f2ab10c5d0270f4b1bdd')

#按照不同的部门导出耗时表
def getHourSpend(depart,excelName,startTime,endTime):
    # time_entries=redmine.time_entry.filter(updated_on='><2016-0-01 00:00:00|2016-06-02 23:59:59',limit=100)
    # time_entries=redmine.time_entry.filter(updated_on=timeScope,limit=20)
    time_entries=redmine.time_entry.all(sort='updated_on:desc')
    i=0
    #time_entry:['activity', 'comments', 'created_on', 'hours', 'id', 'issue', 'project', 'spent_on', 'updated_on', 'user']
    #user:['contacts', 'deals', 'groups', 'id', 'issues', 'memberships', 'name', 'time_entries']
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
        hours=time_entry.hours
        userid=user.id
        curUser=redmine.user.get(userid)
        custom_fields = curUser.custom_fields  #user的自定义字段
        department=custom_fields[0].value[0]    #获得部门名称
        if department==depart:
            if datetime.strptime(startTime,'%Y/%m/%d') <=datetime.strptime(updated_on,'%Y/%m/%d')<=datetime.strptime(endTime,'%Y/%m/%d'):
                i+=1
                # lists = ['项目', '日期', '用户', '活动', '问题', '注释', '耗时', '难易度']
                content = [str(project), updated_on, str(user), str(activity),curIssue, hours,difficultyLevel]
                print(content,depart)
                insertIntoExcel(i,content,excelName)   #结果插入到表格

#获得所有部门耗时，按部门输出,一次性输出所有部门的数据
def exportAllDepartTime():
    #需要导出数据的时间范围，导出不同月份时需要进行想用的修改
    # timeScope='><2016-06-01 00:00:00|2016-06-30 23:59:59'
    startTime='2016/06/01'
    endTime='2016/06/30'
    date='201606'  #excel表格名称的一部分，导出的是哪一个月需要进行修改
    department = [ 'GIS平台部','GIS应用部', '行业服务部', '房地信息事业部', '交通信息事业部',
                   '信息二部', '信息三部','质控部', '规划信息事业部', '暂无部门']
    for depart in department:
        excelName=createExcel(depart,date,PATH)
        #调用查询导出函数
        getHourSpend(depart,excelName,startTime,endTime)
    operateData(date)
exportAllDepartTime()


#每个部门的按人员维度工时填报情况
def operateData(date='201606'):
    exportName=[ 'GIS平台部','GIS应用部', '行业服务部', '房地信息事业部', '交通信息事业部',
                   '信息二部', '信息三部','质控部', '规划信息事业部', '暂无部门']
    for n in range(0,len(exportName)):
        #新建工作簿
        # 新建xls，新建名为sheet1的工作簿
        file = xlwt.Workbook()
        file.add_sheet('log')
        table = file.add_sheet('sheet1')

        lists = ['项目', '日期', '用户', '活动', '主题', '耗时', '难易度']
        # 向新建的sheet1中插入数据
        for i in range(0, len(lists)):
            # 设置列宽
            table.col(i).width = 256 * 16
            # 逐个插入lists列表中的数据，即为表头
            table.write(0, i, lists[i])
        # 设置所要创建的表的名称
        filePath = PATH + '\\redmine_script\\exportResult\\'
        # 保存excel
        file.save(filePath+exportName[n]+'PM系统工时填报'+date+'.xls')
        #插入数据
        #data = xlrd.open_workbook(filePath+exportName[n]+'PM系统工时填报'+date+'.xls')
        #table = data.sheet_by_index(0)

#operateData(date='201606')

# createExcel('test','201606',PATH)


