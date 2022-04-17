"""
 -*- coding: utf-8 -*-

 @Time : 2021/11/5 16:31

 @Author : jagger

 @File : 监考名单生成程序.py

 @Software: PyCharm 

 @contact: 252587809@qq.com

 -*- 功能说明 -*-

"""

import xlwt, xlrd, xlutils
from xlutils.copy import copy
import random

file_name_01 = '2021齐大期末考试监考名单.xls'
file_name_02 = '名单.xls'


def createExcel():
    # 创建表格
    rwb = xlwt.Workbook(encoding='utf-8')
    # cell_overwrite_ok=True 允许重载
    rws = rwb.add_sheet('监考名单', cell_overwrite_ok=True)
    # 合并单元格
    rws.write_merge(0, 1, 0, 5, '2021齐大期末考试监考名单', get_style_01(True))
    rws.write(2, 0, '监考日期', get_style_02())
    rws.write(2, 1, '考试科目', get_style_02())
    rws.write(2, 2, '考试人数', get_style_02())
    rws.write(2, 3, '监考教室', get_style_02())
    rws.write(2, 4, '教室容纳人数', get_style_02())
    rws.write(2, 5, '监考老师', get_style_02())

    # 设置单元格宽度
    rws.col(0).width = 400 * 20
    rws.col(1).width = 350 * 20
    rws.col(2).width = 120 * 20
    rws.col(3).width = 120 * 20
    rws.col(4).width = 200 * 20
    rws.col(5).width = 350 * 20
    # 保存
    rwb.save(file_name_01)


# 写入上午日期
def write_data_morning():
    morning = ['2022/01/03 8:30-11:00', '2022/01/04 8:30-11:00', '2022/01/05 8:30-11:00', '2022/01/06 8:30-11:00',
               '2022/01/07 8:30-11:00', '2022/01/10 8:30-11:00']
    old_excel = xlrd.open_workbook(file_name_01, formatting_info=True)
    new_excel = copy(old_excel)
    rws = new_excel.get_sheet(0)
    t = 0
    for i in range(3, 55, 9):
        for j in range(5):
            rws.write(i + j, 0, morning[t], get_style_02())
        t = t + 1

    new_excel.save(file_name_01)


# 写入下午日期
def write_data_afternoon():
    old_excel = xlrd.open_workbook(file_name_01, formatting_info=True)
    new_excel = copy(old_excel)
    rws = new_excel.get_sheet(0)
    morning = ['2022/01/03 14:30-16:00', '2022/01/04 14:30-16:00', '2022/01/05 14:30-16:00', '2022/01/06 14:30-16:00',
               '2022/01/07 8:30-11:00', '2022/01/10 14:30-16:00']
    t = 0
    for i in range(8, 47, 9):
        for j in range(4):
            rws.write(i + j, 0, morning[t], get_style_02())
        t = t + 1

    for i in range(53, 56):
        rws.write(i, 0, morning[5], get_style_02())
    new_excel.save(file_name_01)


# 写入考试科目以及考试科目人数
def write_items():
    old_excel = xlrd.open_workbook(file_name_01, formatting_info=True)
    new_excel = copy(old_excel)
    rws = new_excel.get_sheet(0)
    j = 3
    for i in range(53):
        rws.write(j, 1, items(file_name_02)[i], get_style_02())
        rws.write(j, 2, items_num(file_name_02)[i], get_style_02())
        j = j + 1
    new_excel.save(file_name_01)


# 写入教室以及教室容纳人数
def write_classroom():
    old_excel = xlrd.open_workbook(file_name_01, formatting_info=True)
    new_excel = copy(old_excel)
    rws = new_excel.get_sheet(0)
    j = 3
    for i in range(53):
        rws.write(j, 3, classroom(file_name_02)[i], get_style_02())
        rws.write(j, 4, classroom_num(file_name_02)[i], get_style_02())
        j = j + 1
    new_excel.save(file_name_01)


# 写入监考教师
def write_teachers():
    old_excel = xlrd.open_workbook(file_name_01, formatting_info=True)
    new_excel = copy(old_excel)
    rws = new_excel.get_sheet(0)
    # (1.3-1.7号)
    get_class_num = xlrd.open_workbook(file_name_01)
    get_num = get_class_num.sheet_by_name('监考名单')
    for i in range(3, 47, 9):
        m_get_data = mTeacher(file_name_02)
        wm_get_data = wmTeacher(file_name_02)
        for j in range(9):
            if get_num.cell_value(i + j, 4) == 30:
                m = random.sample(m_get_data, 1)
                wm = random.sample(wm_get_data, 1)
                people = f"{wm[0]},{m[0]}"
                for rm in m:
                    m_get_data.remove(rm)
                for rwm in wm:
                    wm_get_data.remove(rwm)
                rws.write(i + j, 5, people, get_style_02())
            elif get_num.cell_value(i + j, 4) == 90:
                m = random.sample(m_get_data, 1)
                wm = random.sample(wm_get_data, 2)
                people = ','.join(wm) + f"{','}{m[0]}"
                for rm in m:
                    m_get_data.remove(rm)
                for rwm in wm:
                    wm_get_data.remove(rwm)
                rws.write(i + j, 5, people, get_style_02())
            elif get_num.cell_value(i + j, 4) == 200:
                m = random.sample(m_get_data, 2)
                wm = random.sample(wm_get_data, 2)
                people = f"{','.join(m)},{','.join(wm)}"
                for rm in m:
                    m_get_data.remove(rm)
                for rwm in wm:
                    wm_get_data.remove(rwm)
                rws.write(i + j, 5, people, get_style_02())

    # (1.10)
    m_get_data = mTeacher(file_name_02)
    wm_get_data = wmTeacher(file_name_02)
    for i in range(48, 56):
        if get_num.cell_value(i, 4) == 30:
            m = random.sample(m_get_data, 1)
            wm = random.sample(wm_get_data, 1)
            people = f"{wm[0]},{m[0]}"
            for rm in m:
                m_get_data.remove(rm)
            for rwm in wm:
                wm_get_data.remove(rwm)
            rws.write(i, 5, people, get_style_02())
        elif get_num.cell_value(i, 4) == 90:
            m = random.sample(m_get_data, 1)
            wm = random.sample(wm_get_data, 2)
            people = ','.join(wm) + f"{','}{m[0]}"
            for rm in m:
                m_get_data.remove(rm)
            for rwm in wm:
                wm_get_data.remove(rwm)
            rws.write(i, 5, people, get_style_02())
        elif get_num.cell_value(i, 4) == 200:
            m = random.sample(m_get_data, 2)
            wm = random.sample(wm_get_data, 2)
            people = f"{','.join(m)},{','.join(wm)}"
            for rm in m:
                m_get_data.remove(rm)
            for rwm in wm:
                wm_get_data.remove(rwm)
            rws.write(i, 5, people, get_style_02())

    new_excel.save(file_name_01)


# 考试科目
def items(path):
    item = xlrd.open_workbook(path)
    get_sheet = item.sheet_by_name('考试科目')
    crow = get_sheet.nrows
    row_data = get_sheet.col_values(1, 1, crow)
    return row_data


# 考试人数
def items_num(path):
    item = xlrd.open_workbook(path)
    get_sheet = item.sheet_by_name('考试科目')
    crow = get_sheet.nrows
    row_data = get_sheet.col_values(2, 1, crow)
    return row_data


# 教室
def classroom(path):
    write_classroom = xlrd.open_workbook(path)
    get_sheet_01 = write_classroom.sheet_by_name('空闲教室')
    get_sheet_02 = write_classroom.sheet_by_name('考试科目')
    crow_01 = get_sheet_01.nrows
    crow_02 = get_sheet_02.nrows
    # 教室人数
    class_num = get_sheet_01.col_values(2, 1, crow_01)
    # 考试科目人数
    item_num = get_sheet_02.col_values(2, 1, crow_02)
    i = 0
    j = 0
    row_data = []
    # 53门考试  16个教室
    while i < 53:
        while j < 16:
            # 保证教室能够容纳考试人数
            if (class_num[j] - item_num[i]) == 0 or (class_num[j] - item_num[i]) == 20 or (
                    class_num[j] - item_num[i]) == 30 or (class_num[j] - item_num[i]) == 80 or (
                    class_num[j] - item_num[i]) == 90:
                row_data.append(get_sheet_01.cell_value(j + 1, 1))
                i = i + 1
                j = j + 1
            else:
                j = j + 1
            if i >= 53:
                break
        j = 0
    return row_data


# 教室容纳人数
def classroom_num(path):
    write_classroom = xlrd.open_workbook(path)
    get_sheet_01 = write_classroom.sheet_by_name('空闲教室')
    get_sheet_02 = write_classroom.sheet_by_name('考试科目')
    crow_01 = get_sheet_01.nrows
    crow_02 = get_sheet_02.nrows
    # 教室人数
    class_num = get_sheet_01.col_values(2, 1, crow_01)
    # 考试科目人数
    item_num = get_sheet_02.col_values(2, 1, crow_02)
    i = 0
    j = 0
    row_data = []
    # 53门考试  16个教室
    while i < 53:
        while j < 16:
            if (class_num[j] - item_num[i]) == 0 or (class_num[j] - item_num[i]) == 20 or (
                    class_num[j] - item_num[i]) == 30 or (class_num[j] - item_num[i]) == 80 or (
                    class_num[j] - item_num[i]) == 90:
                row_data.append(get_sheet_01.cell_value(j + 1, 2))
                i = i + 1
                j = j + 1
            else:
                j = j + 1
            if i >= 53:
                break
        j = 0

    return row_data


# 男监考老师
def mTeacher(path):
    write_teacher = xlrd.open_workbook(path)
    get_sheet = write_teacher.sheet_by_name('监考人员名单')
    row_data = get_sheet.col_values(2, 1, 29)
    # row_data = (random.sample(mteacher,i))
    return row_data


# 女监考老师
def wmTeacher(path):
    write_teacher = xlrd.open_workbook(path)
    get_sheet = write_teacher.sheet_by_name('监考人员名单')
    row_data = get_sheet.col_values(2, 30, 74)
    # row_data = (random.sample(wmteacher, i))
    return row_data


# 表格样式1
def get_style_01(bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = "宋体"
    font.bold = bold
    font.underline = False
    font.italic = False
    font.colour_index = 0
    font.height = 300  # 200为10号字体
    style.font = font

    # 单元格居中
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    align.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
    style.alignment = align
    border = xlwt.Borders()  # 给单元格加框线
    border.left = xlwt.Borders.THIN  # 左
    border.top = xlwt.Borders.THIN  # 上
    border.right = xlwt.Borders.THIN  # 右
    border.bottom = xlwt.Borders.THIN  # 下
    style.borders = border
    return style


# 表格样式2
def get_style_02():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = "宋体"
    font.underline = False
    font.italic = False
    font.colour_index = 0
    font.height = 200  # 200为10号字体
    style.font = font

    # 单元格居中
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    align.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
    style.alignment = align
    border = xlwt.Borders()  # 给单元格加框线
    border.left = xlwt.Borders.THIN  # 左
    border.top = xlwt.Borders.THIN  # 上
    border.right = xlwt.Borders.THIN  # 右
    border.bottom = xlwt.Borders.THIN  # 下
    style.borders = border
    return style


if __name__ == "__main__":
    createExcel()
    write_data_morning()
    write_data_afternoon()
    write_items()
    write_classroom()
    write_teachers()

    print("创建成功！")
