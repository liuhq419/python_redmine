# -*- coding:utf-8 -*-
from xlwt import *
FILENAME = 'Project_result.xls'


# font0 = Font()
# font0.name = 'Times New Roman'
# font0.struck_out = True
# font0.bold = True
#
# style0 = XFStyle()
# style0.font = font0

'''标题字体style_title
    有左右下边界，加粗，宋体'''
#设置字体
fnt = Font()
fnt.name = '宋体'
# fnt.colour_index = i
fnt.bold = True
fnt.outline = True
#设置边界
borders = Borders()
borders.left = 1
borders.right = 1
borders.bottom = 1

style_title = XFStyle()
style_title.font = fnt
style_title.borders = borders
'''人员 person_style
    宋体，有边框'''
person_font = Font()
person_font.name = '宋体'

person_borders = Borders()
person_borders.bottom = 1
person_borders.right = 1
person_borders.bottom = 1
person_borders.top = 1

person_style = XFStyle()
person_style.font = person_font
person_style.borders = person_borders


'''项目汇总project_style 汇总均用的该字体
    宋体。有背景色'''
project_font = Font()
project_font.bold = True
project_font.name = '宋体'

project_pattern = Pattern()
project_pattern.pattern = project_pattern.SOLID_PATTERN
project_pattern.pattern_fore_colour = 0x2c

project_style = XFStyle()
project_style.font = project_font
project_style.pattern = project_pattern

'''部门标题department_title_style
    宋体加粗加高'''
dpt_font = Font()
dpt_font.bold = True
dpt_font.name = '宋体'
dpt_font.height = 256*2

dpt_alignment = Alignment()
dpt_alignment.horz = Alignment.HORZ_CENTER
dpt_alignment.vert = Alignment.VERT_CENTER

dpt_title_style = XFStyle()
dpt_title_style.font = dpt_font
dpt_title_style.alignment = dpt_alignment
#行格式
dpt_title_row0 = easyxf('font:height 720;')
'''部门第二行标题department_style
    宋体加粗有边框'''
dpt_font_1 = Font()
dpt_font_1.bold = True
dpt_font_1.name = '宋体'
dpt_font_1.outline = True

dpt_borders = Borders()
dpt_borders.bottom = 1
dpt_borders.right = 1
dpt_borders.bottom = 1
dpt_borders.top = 1

dpt_title_style_1 = XFStyle()
dpt_title_style_1.font = dpt_font_1
dpt_title_style_1.borders = dpt_borders

'''有背景色的普通字体high_style'''
high_pattern = Pattern()
high_pattern.pattern = project_pattern.SOLID_PATTERN
high_pattern.pattern_fore_colour = 0x2b

high_style = XFStyle()
high_style.pattern = high_pattern