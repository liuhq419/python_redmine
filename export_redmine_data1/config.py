# -*- coding:utf-8 -*-
# 配置文件
'''
    起始日期STARTDATE,截止日期ENDDATE
    注：输出结果包含日期当天
    格式：  '2016-01-01'
'''
#以下五个参数是需要进行做相应修改的参数
STARTDATE = '2016-07-01'
ENDDATE = '2016-07-31'
MONTH = 7  # 第MONTH月份的表 int类型：不加引号的数字
date='201607'     #查询年月
workDate=22  #workDate为公司每个月的工作天数，如果没有节假日，则默认为22天，否则要扣节假日时间


'''登陆方式：key
获取方式：右上角（我的账号）——API访问键——显示——双击全选——复制粘贴'''
MYKEY = '885575e983a0fd543048f2ab10c5d0270f4b1bdd'





'''输入文件名
    注意 后缀 .xls
    格式：  'filename.xls '
    '''
# 修改好的项目经理名单
MANAGE_LIST = '经理名单.xls'

''' 输出文件名
    注意后缀 .xls
     格式：  'filename.xls '
     '''
# 项目维度-按项目经理
FILENAME_MANAGE_OUT = '项目经理结果.xls'
# 项目维度-按部门
FILENAME_DEPARTMENT = '部门结果.xls'

# 导出项目包含的项目经理 （需要人工校对excel的内容）
READ_MANAGE = MANAGE_LIST #不需要修改


'''redmine地址（若网址不变，无需修改）
    格式： 'http://pm.dpark.com.cn'
    '''
URL = 'http://pm.dpark.com.cn'


import os
PATH = os.path.split(os.path.realpath(__file__))[0]
# print(PATH)
DATE = '><'+STARTDATE+'|'+ENDDATE
FILENAME_MANAGE_OUT = os.path.join(PATH, 'exportResult', 'PM系统项目完成情况', FILENAME_MANAGE_OUT)
FILENAME_DEPARTMENT = os.path.join(PATH, 'exportResult', 'PM系统项目完成情况', FILENAME_DEPARTMENT)
MANAGE_LIST = os.path.join(PATH, MANAGE_LIST)
READ_MANAGE = os.path.join(PATH, READ_MANAGE)

