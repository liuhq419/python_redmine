# -*- coding:utf-8 -*-
import time
from redmine import Redmine
from public_script.data.excel_data import project_style, high_style
from public_script.Creat_Excel import creat_excel, creat_sheet_department
from public_script.public import all_indexs
from config import DATE, FILENAME_DEPARTMENT, URL, MYKEY


def dpark_department():
    redmine = Redmine(URL, key=MYKEY)

    filename = FILENAME_DEPARTMENT
    wb = creat_excel()
    START_ROW = 2
    # 从第三行开始写入

    issues = redmine.issue.filter(due_date=DATE, status_id='*')

    name_list = []
    project_list = []
    subject_list = []
    due_date_list = []
    status_list = []

    department_list = []

    error_list = []
    has_error = False

    status_new = 0
    status_close = 0
    i = 1 #问题计数
    
    #原始数据
    ws_0 = wb.add_sheet('原始数据')
    ws_0.write(0, 0, '项目名称')
    ws_0.write(0, 1, '主要执行人')
    ws_0.write(0, 2, '所属部门')
    ws_0.write(0, 3, '工作内容')
    ws_0.write(0, 4, '计划完成时间')
    ws_0.write(0, 5, '状态')


    for issues_one in issues:
        project_name = issues_one.project.name
        project_list.append(project_name)
        ws_0.write(i, 0, project_name)
        '''人员'''
        # issues_name = issues_one.assigned_to.name
        try:
            for resource in issues_one.custom_fields.resources:

                if resource.get('name', 'nothing') == '主要执行人': #获取执行人名
                # if resource['name'] == '主要执行人': #获取执行人名
                    name_id = resource['value']
                    name_user = redmine.user.get(name_id)
                    issues_name = name_user.lastname + name_user.firstname
                    name_list.append(issues_name)
                    ws_0.write(i, 1, issues_name)
                    name_department = name_user.custom_fields.resources[0]['value'] #获得执行人所属部门（list）
                    department_list.append(name_department[0])
                    ws_0.write(i, 2, name_department[0])
                    break
        except :
            has_error = True
            t = time.strftime('%Y-%m-%d %X', time.localtime(time.time()))
            ws_0.write(i, 1, 'null')
            f = open(r'error.txt', 'w')
            error = '该问题没有主要执行人，请检查主题:'+ issues_one.subject
            f.write(error)
            f.write('\n报告时间：%s' % t)
            ws_0.write(i, 2, 'null')
            print('该问题没有主要执行人:%s' % issues_one.subject)
            i += 1
            f.close()
            error_list.append(error)
            continue

        '''工作内容'''
        issues_subject = issues_one.subject
        subject_list.append(issues_subject)
        ws_0.write(i, 3, issues_subject)
        '''计划完成时间'''
        issues_due_date = str(issues_one.due_date.year)+'-'+str(issues_one.due_date.month)+'-'+str(issues_one.due_date.day) #or id
        due_date_list.append(issues_due_date)
        ws_0.write(i, 4, issues_due_date)
        '''状态'''
        issues_status = issues_one.status.name
        ws_0.write(i, 5, issues_status)
        if issues_status == '新问题' or issues_status == '已分配':
            status_list.append(1)
        elif issues_status =='已关闭':
            status_list.append(0)
        else:
            error = '存在非‘新问题’也非‘已关闭’也非‘已分配’的问题:' + issues_one.subject
            error_list.append(error)
            raise Exception('存在非‘新问题’也非‘已关闭’也非‘已分配’的问题')
        # print('第%d个问题信息录入完毕' % i)
        i += 1
        

    all_departments = list(set(department_list)) # 提取出全部部门



    for department in all_departments:# 对每个部门执行
        start_row = START_ROW
        sh = creat_sheet_department(department, wb)
        # 查找到该部门下所有信息的序列号
        department_ids = all_indexs(department_list, department)

        # 该部门得到的数据信息
        name_index_by_dpt = []
        pro_index_by_dpt = []
        subject_index_by_dpt = []
        due_date_index_by_dpt = []
        status_index_by_dpt = []
        for department_id in department_ids:
            # 该部门得到的数据信息
            name_index_by_dpt.append(name_list[department_id])
            pro_index_by_dpt.append(project_list[department_id])
            subject_index_by_dpt.append(subject_list[department_id])
            due_date_index_by_dpt.append(due_date_list[department_id])
            status_index_by_dpt.append(status_list[department_id])
        # print('%s信息获取完毕！' % department)

        # 该部门人员名单
        all_names = list(set(name_index_by_dpt))
        for department_name in all_names:
            # 获取一个人的信息id
            name_index_by_name = []
            pro_index_by_name = []
            subject_index_by_name = []
            due_date_index_by_name = []
            status_index_by_name = []
            department_name_ids = all_indexs(name_index_by_dpt, department_name)
            for department_name_id in department_name_ids:
                name_index_by_name.append(name_index_by_dpt[department_name_id])
                pro_index_by_name.append(pro_index_by_dpt[department_name_id])
                subject_index_by_name.append(subject_index_by_dpt[department_name_id])
                due_date_index_by_name.append(due_date_index_by_dpt[department_name_id])
                status_index_by_name.append(status_index_by_dpt[department_name_id])
            # print('%s的问题获取完毕！' % department_name)

            #该人员的所有项目
            all_projects = list(set(pro_index_by_name))

            for his_projects in all_projects:
                # 获取该人的一个项目
                name_index_by_pro = []
                pro_index_by_pro = []
                subject_index_by_pro = []
                due_date_index_by_pro = []
                status_index_by_pro = []

                department_pro_ids = all_indexs(pro_index_by_name, his_projects)
                for department_pro_id in department_pro_ids:
                    name_index_by_pro.append(name_index_by_name[department_pro_id])
                    pro_index_by_pro.append(pro_index_by_name[department_pro_id])
                    subject_index_by_pro.append(subject_index_by_name[department_pro_id])
                    due_date_index_by_pro.append(due_date_index_by_name[department_pro_id])
                    status_index_by_pro.append(status_index_by_name[department_pro_id])
                # print('%s信息获取完毕！' % his_projects)
                '''开始写入！'''
                for write_row in range(0, len(pro_index_by_pro)):


                    if status_index_by_pro[write_row] == 1: #新问题
                        sh.write(start_row+write_row, 4, 1, high_style)
                        status_new += 1
                        sh.write(start_row+write_row, 0, name_index_by_pro[write_row], high_style)
                        sh.write(start_row+write_row, 1, pro_index_by_pro[write_row], high_style)
                        sh.write(start_row+write_row, 2, subject_index_by_pro[write_row], high_style)
                        sh.write(start_row+write_row, 3, due_date_index_by_pro[write_row], high_style)

                    elif status_index_by_pro[write_row] == 0: #已关闭
                        sh.write(start_row+write_row, 5, 1)
                        status_close += 1
                        sh.write(start_row+write_row, 0, name_index_by_pro[write_row])
                        sh.write(start_row+write_row, 1, pro_index_by_pro[write_row])
                        sh.write(start_row+write_row, 2, subject_index_by_pro[write_row])
                        sh.write(start_row+write_row, 3, due_date_index_by_pro[write_row])


                start_row += len(pro_index_by_pro) #录入完一个项目
            '''该人的汇总'''
            sh.write(start_row, 0, department_name, project_style)
            sh.write(start_row, 1, "", project_style)
            sh.write(start_row, 2, "", project_style)
            sh.write(start_row, 3, "", project_style)
            sh.write(start_row, 4, status_new, project_style)
            sh.write(start_row, 5, status_close, project_style)
            sh.write(start_row, 6, status_new+status_close, project_style)
            # 清零
            status_new = 0
            status_close = 0
            start_row += 1

    wb.save(filename)
    print('项目维度-部门\n输出完毕,保存为%s'% filename)
    if has_error:
        print('存在错误:\n')
        for e in error_list:
            print(e)
    return True


if __name__ =='__main__':
    dpark_department()