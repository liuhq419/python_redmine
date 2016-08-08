# -*- coding:utf-8 -*-

from redmine import Redmine
from xlwt import *
from config import *

redmine = Redmine(URL, key = MYKEY)
projects = redmine.project.all(offset=1)
# creat excel
font0 = Font()
font0.name = '宋体'
# font0.struck_out = True
font0.bold = True

style0 = XFStyle()
style0.font = font0

wb = Workbook()
ws0 = wb.add_sheet('0')
ws0.col(1).width = 256*50

ws0.write(0, 0, '项目标识', style0)
ws0.write(0, 1, '项目名称', style0)
ws0.write(0, 2, '项目经理', style0)
ws0.write(0, 3, '只有C列的人名会导入，DEF不用管', style0)
ws0.write(0, 6, '有些项目经理职位的人写的可能不是项目经理', style0)
i = 1

for project in projects:
    manage = []
    ws0.write(i, 0, project.identifier)
    ws0.write(i, 1 , project.name)
    j = 2
    for member in project.memberships:
        for roles in member.roles.resources[0:len(member.roles.resources)]:
            role = roles.get('name')
            if role == '项目经理':
                manage.append(member.user.name)
                # print(member.user.name)
    manage_set = list(set(manage))
    #没有项目经理，或该经理的职位写的不是’项目经理‘
    if len(manage_set) == 0:
        ws0.write(i, j, '请填写项目经理', style0)

    for mange_name in manage_set:
        ws0.write(i, j, mange_name)
        j += 1

    i += 1
wb.save(MANAGE_LIST)