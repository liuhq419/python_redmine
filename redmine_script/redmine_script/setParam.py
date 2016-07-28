# coding:utf-8
from redmine import Redmine

#按照更新时间进行过滤，导出什么月份的数据需要进行相应的修改
startDate='2016-07-01'   #查询的开始日期
endDate='2016-07-31'     #查询的结束日期
date='201607'            #查询月份

#认证到项目管理系统，查看当前账号下的‘key’，输入key的值
key='885575e983a0fd543048f2ab10c5d0270f4b1bdd'


redmine = Redmine('http://pm.dpark.com.cn/', key=key)

filterDate='><'+startDate+'|'+endDate
