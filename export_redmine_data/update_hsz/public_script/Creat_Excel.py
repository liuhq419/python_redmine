# -*- coding:utf-8 -*-
__author__ = 'hanshangzhen'
from xlwt import *

from config import MONTH
from public_script.data.excel_data import *


def creat_excel():
    wb = Workbook()
    # w = wb.add_sheet('sheet1')
    # dir_result = os.path.join(os.path.dirname(os.path.dirname(__file__)), filename)
    # excel.save(dir_result)
    return wb


def creat_sheet(name, wb):
    # dir_result = os.path.join(os.path.dirname(os.path.dirname(__file__)), FILENAME)
    # wb = xlrd.open_workbook(dir_result)
    # excel =  copy(wb)
    wb= wb

    # #列格式
    # font0 = Font()
    # font0.name = '宋体'
    # font0.bold = True
    #
    # style0 = XFStyle()
    # style0.font = font0
    # #行格式
    # borders = Borders()
    # borders.left = 6
    # borders.right = 6
    # borders.top = 6
    # borders.bottom = 6
    #
    # al = Alignment()
    # al.horz = Alignment.HORZ_CENTER
    # al.vert = Alignment.VERT_CENTER



    tall_style = easyxf('font:height 360;')
    # tall_style.font.bold = True
    # tall_style.borders = borders
    # tall_style.alignment.horz = Alignment.HORZ_CENTER
    # tall_style.alignment.vert = Alignment.VERT_CENTER
    ws0 = wb.add_sheet(name)
    #锁定第一行
    ws0.panes_frozen = True
    ws0.horz_split_pos = 1

    ws0.row(0).set_style(tall_style)
    ws0.col(0).width = 256*10
    ws0.col(1).width = 256*30
    ws0.col(2).width = 256*10
    ws0.col(3).width = 256*50
    ws0.col(4).width = 256*15
    ws0.col(5).width = 256*13
    ws0.col(6).width = 256*13
    ws0.col(7).width = 256*13

    ws0.write(0, 0, '项目经理', style_title)
    ws0.write(0, 1, '项目名称', style_title)
    ws0.write(0, 2, '项目人员', style_title)
    ws0.write(0, 3, '工作内容', style_title)
    ws0.write(0, 4, '计划完成时间', style_title)
    ws0.write(0, 5, '已完成任务', style_title)
    ws0.write(0, 6, '未完成任务', style_title)
    ws0.write(0, 7, '汇总', style_title)

    return ws0

def creat_sheet_department(name, wb):
    ws0 = wb.add_sheet(name)
    ws0.panes_frozen = True
    ws0.horz_split_pos = 2
    ws0.row(0).set_style(dpt_title_row0)
    ws0.write_merge(0, 0, 0, 6, '%s%d月计划任务完成情况' % (name, MONTH), dpt_title_style)
    ws0.write(1, 0, '姓名', dpt_title_style_1)
    ws0.write(1, 1, '项目', dpt_title_style_1)
    ws0.write(1, 2, '主题编号', dpt_title_style_1)
    ws0.write(1, 3, '主题', dpt_title_style_1)
    ws0.write(1, 4, '计划完成日期', dpt_title_style_1)
    ws0.write(1, 5, '新问题', dpt_title_style_1)
    ws0.write(1, 6, '已关闭', dpt_title_style_1)
    ws0.write(1, 7, '总计', dpt_title_style_1)
    ws0.col(0).width = 256*15
    ws0.col(1).width = 256*40
    ws0.col(2).width = 256*10
    ws0.col(3).width = 256*60
    ws0.col(4).width = 256*15
    ws0.col(5).width = 256*11
    ws0.col(6).width = 256*11
    ws0.col(7).width = 256*11
    return ws0

def creat_sheet_workplan(name,wb):
    ws0 = wb.add_sheet(name)

    ws0.row(0).set_style(dpt_title_row0)
    # ws0.write_merge(0, 0, 0, 6, '本周工作计划', plan_title_style)
    # ws0.write(1, 0, '项目经理', dpt_title_style_1)
    # ws0.write(1, 1, '项目名称', dpt_title_style_1)
    # ws0.write(1, 2, '项目成员', dpt_title_style_1)
    # ws0.write(1, 3, '工作内容', dpt_title_style_1)
    # ws0.write(1, 4, '优先级', dpt_title_style_1)
    # ws0.write(1, 5, '计划完成日期', dpt_title_style_1)
    # ws0.write(1, 6, '新建问题', dpt_title_style_1)
    ws0.col(0).width = 256*15
    ws0.col(1).width = 256*40
    ws0.col(2).width = 256*20
    ws0.col(3).width = 256*60
    ws0.col(4).width = 256*9
    ws0.col(5).width = 256*8
    ws0.col(6).width = 256 * 8
    ws0.col(7).width = 256*13
    ws0.col(8).width = 256 * 8

    return ws0


def creat_sheet_LastWeekPlan(sh, row):
    sh.write_merge(row, row, 0, 8, '任务回顾表', plan_title_style)
    sh.row(row).set_style(dpt_title_row0)
    row += 1
    sh.write_merge(row, row, 0, 8, '计划内工作（上周一创建的任务）与上上周遗留任务', plan_title_1_style)
    sh.row(row).set_style(plan_title_row)
    row += 1
    sh.write(row, 0, '项目经理', project_style)
    sh.write(row, 1, '项目名称', project_style)
    sh.write(row, 2, '项目成员', project_style)
    sh.write(row, 3, '工作内容', project_style)
    sh.write(row, 4, '任务编号', project_style)
    sh.write(row, 5, '优先级', project_style)
    sh.write(row, 6, '难易度', project_style)
    sh.write(row, 7, '计划完成日期', project_style)
    sh.write(row, 8, '状态', project_style)
    return  sh

if __name__ == '__main__':
    e = creat_excel()
    creat_sheet( 'test',e)
    e.save('test.xls')
