# -*- coding:utf-8 -*-
__author__ = 'hanshangzhen'
from xlwt import *

from config import MONTH
from public_script.data.excel_data import style_title, dpt_title_row0, dpt_title_style, dpt_title_style_1


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
    ws0.write(1, 2, '主题', dpt_title_style_1)
    ws0.write(1, 3, '计划完成日期', dpt_title_style_1)
    ws0.write(1, 4, '新问题', dpt_title_style_1)
    ws0.write(1, 5, '已关闭', dpt_title_style_1)
    ws0.write(1, 6, '总计', dpt_title_style_1)
    ws0.col(0).width = 256*15
    ws0.col(1).width = 256*40
    ws0.col(2).width = 256*60
    ws0.col(3).width = 256*15
    ws0.col(4).width = 256*11
    ws0.col(5).width = 256*11
    ws0.col(6).width = 256*11
    return ws0



if __name__ == '__main__':
    e = creat_excel()
    creat_sheet( 'test',e)
    e.save('test.xls')
