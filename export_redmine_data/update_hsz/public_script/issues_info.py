from redmine import Redmine
import datetime

class Issues_Info(object):
    def __init__(self, redmine):
        self.issues_list = []
        self.name_list = []
        self.redmine = redmine
        self.number = 0

    def addInfo(self, issues_one):
        try:
            for resource in issues_one.custom_fields.resources:
                if resource.get('name', 'nothing') == '主要执行人':  # 获取执行人名
                    name_id = resource['value']
                    name_user = self.redmine.user.get(name_id)
                    issues_name = name_user.lastname + name_user.firstname
                    self.name_list.append(issues_name)
                elif resource.get('name', 'nothing') == '难易度':
                    degree = resource['value']
        except:
            error = '该问题没有主要执行人，请检查主题id:' + issues_one.subject + issues_one.id
            self.error_list.append(error)
            print('该问题没有主要执行人:%s ，请检查主题id' % issues_one.subject)
            # continue
        '''工作内容'''
        issues_subject = issues_one.subject
        '''计划完成时间'''
        issues_due_date = issues_one.due_date
        '''状态'''
        issues_status = issues_one.status.name
        '''优先级'''
        issues_priority = issues_one.priority.name
        '''关闭时间'''
        if hasattr(issues_one, 'closed_on'):
            issues_closed_on = issues_one.closed_on
        else:
            issues_closed_on = 0
        self.issues_list.append(issues_name)
        self.issues_list.append(issues_subject)
        self.issues_list.append(issues_due_date)
        self.issues_list.append(issues_status)
        self.issues_list.append(issues_priority)
        self.issues_list.append(issues_closed_on)
        self.issues_list.append(issues_one.id)
        self.issues_list.append(degree)
        self.number += 1

    def addInfo_Legacy(self, issues_one, date_today):
        if hasattr(issues_one, 'closed_on'):
            if issues_one.closed_on < (date_today - datetime.timedelta(days=7)):
                return 0
        try:
            for resource in issues_one.custom_fields.resources:
                if resource.get('name', 'nothing') == '主要执行人':  # 获取执行人名
                    name_id = resource['value']
                    name_user = self.redmine.user.get(name_id)
                    issues_name = name_user.lastname + name_user.firstname
                    self.name_list.append(issues_name)
                elif resource.get('name', 'nothing') == '难易度':
                    degree = resource['value']
        except:
            error = '该问题没有主要执行人，请检查主题id:' + issues_one.subject + issues_one.id
            self.error_list.append(error)
            print('该问题没有主要执行人:%s ，请检查主题id' % issues_one.subject)
            # continue
        '''工作内容'''
        issues_subject = issues_one.subject + '(遗留任务)'
        '''计划完成时间'''
        issues_due_date = issues_one.due_date
        '''状态'''
        issues_status = issues_one.status.name
        '''优先级'''
        issues_priority = issues_one.priority.name
        '''关闭时间'''
        if hasattr(issues_one, 'closed_on'):
            issues_closed_on = issues_one.closed_on
        else:
            issues_closed_on = 0
        self.issues_list.append(issues_name)
        self.issues_list.append(issues_subject)
        self.issues_list.append(issues_due_date)
        self.issues_list.append(issues_status)
        self.issues_list.append(issues_priority)
        self.issues_list.append(issues_closed_on)
        self.issues_list.append(issues_one.id)
        self.issues_list.append(degree)

        self.number += 1


    def getIssuesList(self):
        return self.issues_list

    def getIdList(self):
        return self.issues_id

    def getNameList(self):
        return set(self.name_list)

    def getNumber(self):
        return self.number