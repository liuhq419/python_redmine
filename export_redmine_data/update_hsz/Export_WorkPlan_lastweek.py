# -*- coding: utf-8 -*-
__author__ = 'hanshangzhen'
from redmine import Redmine
from redmine.exceptions import ForbiddenError
import datetime, time, os
from public_script.Creat_Excel import *
from public_script.public import *
from config import MANAGE_LIST, MYKEY, URL, TuesdayDATE, PATH
from public_script.issues_info import Issues_Info

redmine = Redmine(URL, key=MYKEY)
print('开始运行')
filename = MANAGE_LIST
date = open_workbook(filename)
projects_info = date.sheet_by_index(0)
wb = Workbook()
managelist = getManage(projects_info)
#今天日期

date_today = datetime.datetime.now()
if TuesdayDATE:
    date_today = time.strptime(TuesdayDATE, '%Y-%m-%d')
    y, m, d = date_today[0:3]
    date_today = datetime.datetime(y, m, d)

StrToday = date_today.strftime('%Y-%m-%d')
lastMonday = (date_today - datetime.timedelta(days=8)).strftime('%Y-%m-%d')
lastweek_start = (date_today - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
lastweek_end = (date_today - datetime.timedelta(days=2)).strftime('%Y-%m-%d')
lastweek_date = '><' + lastweek_start + '|' + lastweek_end
lastMonday_date = '><' + lastMonday + '|' + lastMonday
legacy_date = getLegacydate(date_today)
error_list = []
i = 0
for manage in managelist:
    i += 1
    print('对项目经理搜索中...%d of %d' % (i, len(managelist)))
    #对每个项目经理进行
    '''获取项目经理名下的项目名单'''
    project_list = getPro(projects_info, manage)
    '''creat sheet'''
    # sh = creat_sheet_workplan(manage,wb)
    sh = creat_sheet_workplan(manage,wb) #包括本周工作计划表头
    writeRow = 0

    ''' 创建本周工作计划'''

    '''创建上周工作计划'''
    sh = creat_sheet_LastWeekPlan(sh, writeRow)
    # writeRow += 3 # 空行
    # 创建表头,初始化数据
    writeRow += 3
    writeRow_ProStart = writeRow
    writeRow_ManageStart = writeRow
    issues_num = 0
    issues_id = []
    '''复制代码'''
    for pro_id in project_list:
        issues = redmine.issue.filter(project_id=pro_id, created_on=lastMonday, status_id='*')
        pro_name = redmine.project.get(pro_id).name
        # print(pro_name)
        lastweekinfo = Issues_Info(redmine)
        '''添加上周一计划任务'''
        try:
            for issues_one in issues:
                if issues_one.id not in issues_id:
                    lastweekinfo.addInfo(issues_one)
                    issues_id.append(issues_one.id)
        except ForbiddenError:
            error = '无法获取该项目问题信息：' + pro_name
            error_list.append(error)
            print("无法获取问题信息:%s" % pro_name)
            continue

        except:
            error = '发生错误，请检查' + pro_name
            error_list.append(error)
            raise Exception('导出该项目时发生错误%s' % pro_name)

            # 写入问题原始数据
        '''添加遗留任务'''
        issues = redmine.issue.filter(project_id=pro_id, created_on=legacy_date, status_id='*')
        try:
            for issues_one in issues:
                if issues_one.id not in issues_id:
                    lastweekinfo.addInfo_Legacy(issues_one, date_today)
                    issues_id.append(issues_one.id)
        except ForbiddenError:
            error = '无法获取该项目问题信息：' + pro_name
            error_list.append(error)
            print("无法获取问题信息:%s" % pro_name)
            continue

        except:
            error = '发生错误，请检查' + pro_name
            error_list.append(error)
            raise Exception('导出该项目时发生错误%s' % pro_name)
            print(issues_list)

        issues_list = lastweekinfo.getIssuesList()

        if issues_list:
            name_list = lastweekinfo.getNameList()
            for project_person in name_list:
                # 遍历项目人员
                for person_id in range(0, len(issues_list), 8):
                    # 遍历该成员的所有问题
                    # print(person_id)
                    # print(issues_list[person_id])
                    # print(project_person)
                    if issues_list[person_id] == project_person:
                        # 若找到问题，写入
                        # 项目人员
                        closeDate = issues_list[person_id + 5]
                        dueDate = issues_list[person_id + 2]
                        lastweek_style = person_style
                        if closeDate != 0:
                            closeDate = closeDate.date()
                            if closeDate > dueDate:
                                lastweek_style = person_style_Red

                        if issues_list[person_id + 3] == '新问题' or issues_list[person_id + 3] == '已分配':
                            sh.write(writeRow, 8, '未完成', person_style_Red)
                            # sh.write(writeRow, 5, 0, person_style)

                        elif issues_list[person_id + 3] == '已关闭':
                            sh.write(writeRow, 8, '已完成', lastweek_style)

                        else:
                            error_list.append('存在既非已关闭也非新问题也非已分配的问题，查看问题：' + issues_list[person_id])
                            continue

                        sh.write(writeRow, 2, issues_list[person_id], lastweek_style)
                        # 工作内容
                        sh.write(writeRow, 3, issues_list[person_id + 1], lastweek_style)
                        # 计划完成时间
                        sh.write(writeRow, 7, issues_list[person_id + 2].strftime('%Y-%m-%d'), lastweek_style)
                        # 优先级
                        if issues_list[person_id + 4] == '高' or issues_list[person_id + 4] == '非常高' or issues_list[person_id + 4] == '致命':
                            lastweek_style = person_style_Red
                        sh.write(writeRow, 5, issues_list[person_id + 4], lastweek_style)
                        sh.write(writeRow, 4, issues_list[person_id + 6], lastweek_style)
                        sh.write(writeRow, 6, issues_list[person_id + 7], lastweek_style)
                        issues_num += 1

                        # else:
                        #     error = '存在非‘新问题’也非‘已关闭’也非‘已分配’的问题:' + issues_list[person_id + 1]
                        #     error_list.append(error)
                        #     sh.write(writeRow, 8, '不是新问题也不是已关闭也不是已分配')
                        # 该人员的一个问题写入完毕
                        writeRow += 1
                        # print('人员写入完毕！')

            # 项目问题写入完毕
            '''汇总'''
            # sh.write(writeRow, 0, manage, project_style) #项目经理名称
            # sh.write(writeRow, 1, pro_name, project_style)#项目名称
            sh.write_merge(writeRow_ProStart, writeRow, 1, 1, pro_name, word_style_1)
            sh.write_merge(writeRow, writeRow, 2, 7, '汇总', project_style)

            sh.write(writeRow, 8, lastweekinfo.getNumber(), project_style)
            writeRow += 1
            writeRow_ProStart = writeRow
    if writeRow != writeRow_ManageStart:
        writeRow -= 1
        sh.write_merge(writeRow_ManageStart, writeRow, 0, 0, manage, word_style_1)
        writeRow += 1
        sh.write_merge(writeRow, writeRow, 0, 7, '汇总', word_style_1)
        sh.write(writeRow, 8, issues_num, word_style_1)


    '''创建上周额外工作计划'''
    writeRow += 1
    #创建表头
    sh.write_merge(writeRow, writeRow, 0, 8, '计划外工作完成情况（上周二-上周五创建的工作）', plan_title_1_style)
    sh.row(writeRow).set_style(plan_title_row)
    writeRow += 1
    '''复制代码'''
    #初始化数据
    writeRow_ProStart = writeRow
    writeRow_ManageStart = writeRow
    issues_num = 0
    issues_id = []
    for pro_id in project_list:

        issues = redmine.issue.filter(project_id=pro_id, created_on=lastweek_date, status_id='*')
        pro_name = redmine.project.get(pro_id).name
        lastweekAddInfo = Issues_Info(redmine)

        try:
            for issues_one in issues:
                if issues_one.id not in issues_id:
                    lastweekAddInfo.addInfo(issues_one)
                    issues_id.append(issues_one.id)

        except ForbiddenError:
            error = '无法获取该项目问题信息：' + pro_name
            error_list.append(error)
            print("无法获取问题信息:%s" % pro_name)
            continue

        except:
            error = '发生错误，请检查' + pro_name
            error_list.append(error)
            raise Exception('导出该项目时发生错误%s' % pro_name)
            print(issues_list)
            # 写入问题原始数据
        issues_list = lastweekAddInfo.getIssuesList()
        if issues_list:
            name_list = lastweekAddInfo.getNameList()
            for project_person in name_list:
                # 遍历项目人员
                for person_id in range(0, len(issues_list), 8):
                    # 遍历该成员的所有问题
                    # print(person_id)
                    # print(issues_list[person_id])
                    # print(project_person)
                    if issues_list[person_id] == project_person:
                        # 若找到问题，写入
                        # 项目人员
                        closeDate = issues_list[person_id + 5]
                        dueDate = issues_list[person_id + 2]
                        lastweek_style = person_style
                        if closeDate != 0:
                            closeDate = closeDate.date()
                            if closeDate > dueDate:
                                lastweek_style = person_style_Red
                        if issues_list[person_id + 3] == '新问题' or issues_list[person_id + 3] == '已分配':
                            sh.write(writeRow, 8, '未完成', person_style_Red)
                            # sh.write(writeRow, 5, 0, person_style)
                            status_new += 1
                        elif issues_list[person_id + 3] == '已关闭':
                            sh.write(writeRow, 8, '已完成', lastweek_style)
                            status_close += 1
                        else:
                            error_list.append('存在既非已关闭也非新问题也非已分配的问题，查看问题：' + issues_list[person_id])
                            continue

                        sh.write(writeRow, 2, issues_list[person_id], lastweek_style)
                        # 工作内容
                        sh.write(writeRow, 3, issues_list[person_id + 1], lastweek_style)
                        # 计划完成时间
                        sh.write(writeRow, 7, issues_list[person_id + 2].strftime('%Y-%m-%d'), lastweek_style)
                        # 优先级
                        if issues_list[person_id + 4] == '高' or issues_list[person_id + 4] == '非常高' or issues_list[person_id + 4] == '致命':
                            lastweek_style = person_style_Red
                        sh.write(writeRow, 5, issues_list[person_id + 4], lastweek_style)
                        sh.write(writeRow, 4, issues_list[person_id + 6], lastweek_style)
                        sh.write(writeRow, 6, issues_list[person_id + 7], lastweek_style)
                        issues_num += 1
                        writeRow += 1
                        # print('人员写入完毕！')

            # 项目问题写入完毕
            '''汇总'''

            sh.write_merge(writeRow_ProStart, writeRow, 1, 1, pro_name, word_style_1)
            sh.write_merge(writeRow, writeRow, 2, 7, '汇总', project_style)
            sh.write(writeRow, 8, lastweekAddInfo.getNumber(), project_style)
            writeRow += 1
            writeRow_ProStart = writeRow
    if writeRow != writeRow_ManageStart:
        writeRow -= 1
        sh.write_merge(writeRow_ManageStart, writeRow, 0, 0, manage, word_style_1)
        writeRow += 1
        sh.write_merge(writeRow, writeRow, 0, 7, '汇总', word_style_1)
        sh.write(writeRow, 8, issues_num, word_style_1)
    '''备注'''
    writeRow += 1
    sh.write(writeRow, 0, '备注', word_style_1)
    sh.write_merge(writeRow, writeRow, 1, 8, '《上周工作达成表》中，已完成任务标红是由于关闭问题日期超过计划完成日期。而非填报工时日期超过计划完成日期。', person_style)
if error_list:
    print(error_list)
# 保存文件
str = StrToday + '工作计划_re.xls'
file = os.path.join(PATH, 'exportResult', '上周回顾')
if not os.path.exists(file):
    os.makedirs(file)
file = os.path.join(file, str)
wb.save(file)
print('运行完毕，保存路径%s' % file)