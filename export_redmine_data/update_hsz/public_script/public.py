# -*- coding:utf-8 -*-
import redmine, datetime
from xlrd import *
def all_indexs(lst, obj):
    '''
    返回所有obj的index
    '''
    def find_index(lst, obj, start=0):
        try:
            index = lst.index(obj, start)
        except:
            index = -1
        return index
 
    indexes = []
    i = 0
    while True:
        idx = find_index(lst, obj, i)
        if idx == -1:
            return indexes
        indexes.append(idx)
        i = idx + 1
    return indexes


def getManage(projects_info):
    managelist = []
    for i in range(1, projects_info.nrows):
        name = projects_info.cell_value(i, 2)
        managelist.append(name)
    return set(managelist)


def getPro(projects_info, manage):
    projectlist = []
    for i in range(1, projects_info.nrows):
        if manage == projects_info.cell_value(i, 2):
            projectlist.append(projects_info.cell_value(i, 0))
    return projectlist[::-1]

def getLegacydate(date_today):
    '''得到遗留任务日期'''
    lastweek_start = (date_today - datetime.timedelta(days=15)).strftime('%Y-%m-%d')
    lastweek_end = (date_today - datetime.timedelta(days=9)).strftime('%Y-%m-%d')
    lastweek_date = '><' + lastweek_start + '|' + lastweek_end
    return lastweek_date

