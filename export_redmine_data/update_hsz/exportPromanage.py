# -*- coding: utf-8 -*-
__author__ = 'hanshangzhen'
from redmine import Redmine
from redmine.exceptions import ForbiddenError
from xlrd import *

from public_script.data.excel_data import project_style, person_style
from public_script.Creat_Excel import creat_excel, creat_sheet
from config import DATE, FILENAME_MANAGE_OUT, MANAGE_LIST, MYKEY, URL


def dpark_promanage():
    redmine = Redmine(URL, key=MYKEY)
    print('开始运行')
    filename = MANAGE_LIST
    date = open_workbook(filename)
    projects_info = date.sheet_by_index(0)
    # STARTDATE = '2016-07-01'
    # ENDDATE = '2016-07-20'
    # DATE = '><'+STARTDATE+'|'+ENDDATE
    # issues = redmine.issue.all()
    # projects = redmine.project.all(offset=1)
    # projects =redmine.project.get('its-phase1-stage2')


    wb = creat_excel()

    manage_list = []
    sheet_list = []
    manage_sheet = {}
    error_list = []
    issues_id = []
    #原始数据
    ws_0 = wb.add_sheet('原始数据')
    ws_0.write(0, 0, '项目名称')
    ws_0.write(0, 1, '项目经理')
    ws_0.write(0, 2, '项目人员')
    ws_0.write(0, 3, '工作内容')
    ws_0.write(0, 4, '计划完成时间')
    ws_0.write(0, 5, '状态')
    row = 1
    project_id = range(1, projects_info.nrows)
    project_id = reversed(project_id)
    for i in project_id:
        print('项目搜索中：%d of %d' % (projects_info.nrows - i, projects_info.nrows))
        pro_id = projects_info.cell_value(i, 0)
        pro_name = projects_info.cell_value(i, 1)

        issues = redmine.issue.filter(project_id=pro_id, due_date=DATE, status_id='*')
        # print(issues.resources)
        # for issue in issues:
        # if issues.resources!=None:
        '''寻找项目经理'''

        try:
            # print(pro_name)
            issues_list = []
            name_list = []
            status_new = 0
            status_close = 0

            for issues_one in issues:
                person_info = []
                #若该问题未被记录
                if issues_one.id not in issues_id:
                    '''项目人员'''
                    # issues_name = issues_one.assigned_to.name
                    ws_0.write(row, 0, pro_name)
                    ws_0.write(row, 1, projects_info.cell_value(i, 2))
                    try:
                        for resource in issues_one.custom_fields.resources:
                            if resource.get('name', 'nothing') == '主要执行人': #获取执行人名
                                name_id = resource['value']
                                name_user = redmine.user.get(name_id)
                                issues_name = name_user.lastname + name_user.firstname
                                name_list.append(issues_name)
                                ws_0.write(row, 2, issues_name)
                    except:
                        error = '该问题没有主要执行人，请检查主题:'+ issues_one.subject
                        error_list.append(error)
                        ws_0.write(row, 2, 'null')
                        print('该问题没有主要执行人:%s ，请检查主题：' % issues_one.subject)
                        row += 1
                        continue

                    '''工作内容'''
                    issues_subject = issues_one.subject
                    '''计划完成时间'''
                    issues_due_date = str(issues_one.due_date.year)+'-'+str(issues_one.due_date.month)+'-'+str(issues_one.due_date.day) #or id
                    '''状态'''
                    issues_status = issues_one.status.name

                    issues_list.append(issues_name)
                    issues_list.append(issues_subject)
                    issues_list.append(issues_due_date)
                    issues_list.append(issues_status)
                    #写入问题原始数据

                    ws_0.write(row, 3, issues_subject)
                    ws_0.write(row, 4, issues_due_date)
                    ws_0.write(row, 5, issues_status)
                    row += 1
                    #记录问题编号
                    issues_id.append(issues_one.id)

            '''issues_list:
                    0       1       2           3
                项目人员 工作内容 计划完成时间 状态'''
            # print('该项目下问题遍历完毕：')
            # print(issues_list)
            name_list=list(set(name_list))

            if len(name_list) != 0 and issues.total_count != 0:
                manage = projects_info.cell_value(i, 2)

                #是否该项目经理的sheet已经创建
                if manage not in manage_list:
                    manage_list.append(manage)
                    sheet = creat_sheet(manage, wb)
                    manage_sheet[manage] = len(sheet_list)
                    sheet_list.append(sheet)



                sh = sheet_list[manage_sheet[manage]]
                row_start = len(sh.rows)
                person_start = row_start
                #根据项目人员名写入
                # print(name_list)
                for project_person in name_list:
                    #遍历项目人员
                    for person_id in range(0, len(issues_list), 4):
                        #遍历该成员的所有问题
                        # print(person_id)
                        # print(issues_list[person_id])
                        # print(project_person)
                        if issues_list[person_id] == project_person:
                            #若找到问题，写入
                            #项目人员
                            # print('找到问题！')
                            sh.write(person_start, 2, issues_list[person_id], person_style)
                            #工作内容
                            sh.write(person_start, 3, issues_list[person_id+1], person_style)
                            #计划完成时间
                            sh.write(person_start, 4, issues_list[person_id+2], person_style)

                            #任务状态
                            if issues_list[person_id+3] == '已关闭':
                                sh.write(person_start, 5, 1, person_style)
                                sh.write(person_start, 6, 0, person_style)
                                status_close += 1
                            elif issues_list[person_id+3] == '新问题' or issues_list[person_id+3] == '已分配':
                                sh.write(person_start, 6, 1, person_style)
                                sh.write(person_start, 5, 0, person_style)
                                status_new += 1
                            else:
                                error = '存在非‘新问题’也非‘已关闭’也非‘已分配’的问题:' + issues_list[person_id+1]
                                error_list.append(error)
                                sh.write(person_start, 8, '不是新问题也不是已关闭也不是已分配')
                            #该人员的一个问题写入完毕
                            person_start += 1
                            # print('人员写入完毕！')
                #项目问题写入完毕
                '''汇总'''
                sh.write(person_start, 0, manage, project_style) #项目经理名称
                sh.write(person_start, 1, pro_name, project_style)#项目名称
                sh.write(person_start, 2, '', project_style)
                sh.write(person_start, 3, '', project_style)
                sh.write(person_start, 4, '', project_style)
                sh.write(person_start, 5, status_close, project_style)#已完成任务
                sh.write(person_start, 6, status_new, project_style)#未完成任务
                sh.write(person_start, 7, status_close+status_new, project_style)#汇总
                #保存
                # print('项目写入完毕！')

        except ForbiddenError:
            error = '无法获取该项目问题信息：'+ pro_name
            error_list.append(error)
            print("无法获取问题信息:%s" % pro_name)
            pass
        except:
            error = '发生错误，请检查' + pro_name
            error_list.append(error)
            raise Exception('导出该项目时发生错误%s' % pro_name)

    # i=0
    if len(error_list) != 0:
        print('存在错误:\n')
        for e in error_list:
            print(e)
    wb.save(FILENAME_MANAGE_OUT)
    print('项目维度-经理\n输出完毕,保存为%s'% FILENAME_MANAGE_OUT)
    return True

if __name__ =='__main__':
    dpark_promanage()