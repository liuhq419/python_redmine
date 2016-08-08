# -*- coding:utf-8 -*-
from xlwt import *
FILENAME = 'Project_result.xls'

#边框 Borders.
borders_0 = Borders()
borders_0.bottom = 1
borders_0.right = 1
borders_0.bottom = 1
borders_0.top = 1

'''标题字体style_title
    有左右下边界，加粗，Arial Unicode MS'''
#设置字体
fnt = Font()
fnt.name = 'Arial Unicode MS'
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
    Arial Unicode MS，有边框'''
person_font = Font()
person_font.name = 'Arial Unicode MS'

person_borders = Borders()
person_borders.bottom = 1
person_borders.right = 1
person_borders.bottom = 1
person_borders.top = 1

person_style = XFStyle()
person_style.font = person_font
person_style.borders = person_borders

'''人员 person_style_red
    Arial Unicode MS，有边框，红色'''
person_font_Red = Font()
person_font_Red.name = 'Arial Unicode MS'
person_font_Red.colour_index = 0x2

person_style_Red = XFStyle()
person_style_Red.borders = person_borders
person_style_Red.font = person_font_Red

'''项目汇总project_style 汇总均用的该字体
    Arial Unicode MS。有背景色'''
project_font = Font()
project_font.bold = True
project_font.name = 'Arial Unicode MS'

project_pattern = Pattern()
project_pattern.pattern = project_pattern.SOLID_PATTERN
project_pattern.pattern_fore_colour = 0x2c

project_style = XFStyle()
project_style.font = project_font
project_style.pattern = project_pattern

'''部门标题department_title_style
    Arial Unicode MS加粗加高'''
dpt_font = Font()
dpt_font.bold = True
dpt_font.name = 'Arial Unicode MS'
dpt_font.height = 256*2

dpt_alignment = Alignment()
dpt_alignment.horz = Alignment.HORZ_CENTER
dpt_alignment.vert = Alignment.VERT_CENTER

dpt_title_style = XFStyle()
dpt_title_style.font = dpt_font
dpt_title_style.alignment = dpt_alignment
'''工作计划标题plan_title_style
    Arial Unicode MS加粗加高'''
plan_font = Font()
plan_font.bold = True
plan_font.name = 'Arial Unicode MS'
plan_font.height = 256*1
plan_font.outline = True

plan_alignment = Alignment()
plan_alignment.horz = Alignment.HORZ_CENTER
plan_alignment.vert = Alignment.VERT_CENTER

plan_title_style = XFStyle()
plan_title_style.font = plan_font
plan_title_style.alignment = plan_alignment
plan_title_style.borders = borders_0



'''工作计划小标题 plan_title_1_style
    Arial Unicode MS加粗居中 有边框'''
plan_font1 = Font()
plan_font1.bold = True
plan_font1.name = 'Arial Unicode MS'
plan_font1.outline = True


plan_title_1_style = XFStyle()
plan_title_1_style.font = plan_font1
plan_title_1_style.alignment = plan_alignment
plan_title_1_style.borders= borders_0
'''部门第二行标题department_style
    Arial Unicode MS加粗有边框'''
dpt_font_1 = Font()
dpt_font_1.bold = True
dpt_font_1.name = 'Arial Unicode MS'
dpt_font_1.outline = True


dpt_title_style_1 = XFStyle()
dpt_title_style_1.font = dpt_font_1
dpt_title_style_1.borders = borders_0

'''有背景色的普通字体high_style'''
high_pattern = Pattern()
high_pattern.pattern = project_pattern.SOLID_PATTERN
high_pattern.pattern_fore_colour = 0x2b

high_style = XFStyle()
high_style.pattern = high_pattern

'''普通字体 word_style_1
    Arial Unicode MS 有边框 加粗 居中'''
word_font_1 = Font()
word_font_1.bold = True
word_font_1.name = 'Arial Unicode MS'
word_font_1.outline = True

word_borders_1 = Borders()
word_borders_1.bottom = 1
word_borders_1.right = 1
word_borders_1.bottom = 1
word_borders_1.top = 1

word_alignment_1 = Alignment()
word_alignment_1.vert = Alignment.VERT_CENTER

word_style_1 = XFStyle()
word_style_1.font = word_font_1
word_style_1.borders = word_borders_1
word_style_1.alignment = word_alignment_1

'''行格式'''
dpt_title_row0 = easyxf('font:height 720;')
'''行格式2'''
plan_title_row = easyxf('font:height 500;')