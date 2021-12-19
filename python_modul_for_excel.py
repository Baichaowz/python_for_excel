# -*- coding: utf-8 -*-
# @Time     : 2021-11-16 16:53
# @Author   : Chardman
# @Software : Pycharm

# 1.xlrd模块
# xlrd模块可以用于读取Excel的数据，速度非常快，推荐使用！(官方文档：https://xlrd.readthedocs.io/en/latest/)
# xlrq可以读取xls、xlsx文件
# 1.1 open_workbook
from xlrd import open_workbook

wb1 = open_workbook('excel-file.xlsx')

# 1.2 获取一个表sheet
table1 = wb1.sheets()[0]
table2 = wb1.sheet_by_index(0)
table3 = wb1.sheet_by_name('Sheet1')
# 以上三个函数都会返回一个xlrd.sheet.Sheet()对象

names = wb1.sheet_names()  # 返回book中所有工作表的名字
wb1.sheet_loaded('Sheet1')  # 检查某个sheet是否导入完毕

# 1.3 行的操作
num_rows = table1.nrows
# 获取该sheet中的行数，注，这里table.nrows后面不带().
table1.row(0)
# 返回由该行中所有的单元格对象组成的列表。
table1.row_slice(0)
table1.row_slice(rowx=2, start_colx=1, end_colx=None)
# 返回由该行中特定范围的单元格对象组成的列表
table1.row_types(rowx=0, start_colx=1, end_colx=None)
# 返回由该行中所有单元格的数据类型组成的列表；
# 返回值为逻辑值列表，若类型为empty则为0，否则为1
table1.row_values(0, start_colx=0, end_colx=None)
# 返回由该行中所有单元格的数据组成的列表
table1.row_len(0)
# 返回该行的有效单元格长度，即这一行有多少个数据

# 1.4 列的操作
num_cols = table1.ncols
# 获取列表的有效列数
table1.col(0, start_rowx=0, end_rowx=None)
# 返回由该列中所有的单元格对象组成的列表
table1.col_slice(1, start_rowx=0, end_rowx=None)
# 返回由该列中所有的单元格对象组成的列表
table1.col_types(1, start_rowx=0, end_rowx=None)
# 返回由该列中所有单元格的数据类型组成的列表
table1.col_values(0, start_rowx=0, end_rowx=None)
# 返回由该列中所有单元格的数据组成的列表

# 1.5 单元格的操作
table1.cell(0, 0)
# 返回单元格对象
table1.cell_type(0, 0)
# 返回对应位置单元格中的数据类型
table1.cell_value(0, 0)
# 返回对应位置单元格中的数据

# 1.6 迭代
row1 = table1.row(0)
for each_cell in row1:
    print(each_cell, each_cell.value, type(each_cell))

# 2.xlwt模块
# 仅限于xls格式
# xlwt可以用于写入新的Excel表格或者在原表格基础上进行修改，速度也很快，推荐使用！
# 3.2.2 使用xlwt创建新表格并写入
import xlwt


# 2.1 写入新表格
def excel_writer1():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")
    worksheet.write(2, 1, "内容2")

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer1()


# 2.1 设置字体格式
def excel_writer2():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 初始化样式
    style = xlwt.XFStyle()

    # 为样式创建字体
    font = xlwt.Font()
    font.name = 'Times New Roman'  # 字体
    font.bold = True  # 加粗
    font.underline = True  # 下划线
    font.italic = True  # 斜体

    # 设置样式
    style.font = font

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")
    worksheet.write(2, 1, "内容2", style)

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer2()


# 2.2 设置列宽
# xlwt中列宽的值表示方法：默认字体0的1/256为衡量单位。
# xlwt创建时使用的默认宽度为2960，既11个字符0的宽度
# 所以我们在设置列宽时可以用如下方法：
# width = 256 * 20 256为衡量单位，20表示20个字符宽度
def excel_writer3():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 初始化样式
    style = xlwt.XFStyle()

    # 为样式创建字体
    font = xlwt.Font()
    font.name = 'Times New Roman'  # 字体
    font.bold = True  # 加粗
    font.underline = True  # 下划线
    font.italic = True  # 斜体

    # 设置样式
    style.font = font

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")
    worksheet.write(2, 1, "内容2", style)

    # 设置列宽
    worksheet.col(0).width = 256 * 20

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer3()


# 2.3 xlwt 设置行高
# 在xlwt中没有特定的函数来设置默认的列宽及行高
# 行高是在单元格的样式中设置的，你可以通过自动换行、输入文字的多少来确定行高
def excel_writer4():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 初始化样式
    style = xlwt.XFStyle()

    # 为样式创建字体
    font = xlwt.Font()
    font.name = 'Times New Roman'  # 字体
    font.bold = True  # 加粗
    font.underline = True  # 下划线
    font.italic = True  # 斜体

    # 设置样式
    style.font = font

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")
    worksheet.write(2, 1, "内容2", style)

    # 设置列宽
    worksheet.col(0).width = 256 * 20

    # 设置行高
    style1 = xlwt.easyxf('font: height 360;')
    row = worksheet.row(0)
    row.set_style(style1)

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer4()


# 2.4 xlwt合并列和行
def excel_writer5():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")

    # 合并 第1行到第2行 的 第0列到第3列
    worksheet.write_merge(1, 2, 0, 3, 'Merge Test')

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer5()


# 2.5 xlwt添加边框
def excel_writer6():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")

    # 设置边框样式
    borders = xlwt.Borders()  # Create Borders

    # May be:   NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR,
    #           MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED,
    #           MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    # DASHED虚线
    # NO_LINE没有
    # THIN实线

    borders.left = xlwt.Borders.DASHED
    borders.right = xlwt.Borders.DASHED
    borders.top = xlwt.Borders.DASHED
    borders.bottom = xlwt.Borders.DASHED
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40

    style = xlwt.XFStyle()  # Create Style
    style.borders = borders  # Add Borders to Style

    worksheet.write(2, 1, "内容2", style)

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer6()


# 2.6 xlwt设置背景颜色
def excel_writer7():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")

    # 创建样式
    pattern = xlwt.Pattern()

    # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN

    # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow,
    # 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow ,
    # almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    pattern.pattern_fore_colour = 5
    style = xlwt.XFStyle()
    style.pattern = pattern

    # 使用样式
    worksheet.write(2, 1, "内容2", style)

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer7()


# 2.7 xlwt设置单元格对齐
def excel_writer8():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    # 往表格写入内容
    worksheet.write(0, 0, "内容1")

    # 设置样式
    style = xlwt.XFStyle()
    al = xlwt.Alignment()
    # VERT_TOP = 0x00       上端对齐
    # VERT_CENTER = 0x01    居中对齐（垂直方向上）
    # VERT_BOTTOM = 0x02    低端对齐
    # HORZ_LEFT = 0x01      左端对齐
    # HORZ_CENTER = 0x02    居中对齐（水平方向上）
    # HORZ_RIGHT = 0x03     右端对齐
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al

    # 对齐写入
    worksheet.write(2, 1, "内容2", style)

    # 保存
    workbook.save("new_excel_file.xls")


excel_writer8()


# 2.8 xlwt设置输出格式汇总
def excel_writer():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')
    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = u'楷体'
    font.bold = True
    font.underline = True
    font.italic = True

    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40

    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern_fore_colour = 5

    alignment = xlwt.Alignment()
    alignment.horz = 0x02  # 设置水平居中
    alignment.vert = 0x01  # 设置垂直居中

    style.font = font
    style.borders = borders
    style.pattern = pattern
    style.alignment = alignment

    # 简单写入
    worksheet.write(0, 0, "内容1", style)

    # 合并写入
    worksheet.write_merge(1, 2, 0, 3, 'Merge Test', style)

    workbook.save('xlwt写入操作汇总.xls')


excel_writer()


# 将xlwt的所有颜色输出到excel
def colors_collect():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')
    # 创建新的sheet表
    worksheet = workbook.add_sheet("Sheet1")

    for i in range(256):
        style = xlwt.XFStyle()
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = i
        style.pattern = pattern
        worksheet.write(i, 0, i, style)
    workbook.save('xlwt背景颜色汇总.xls')


colors_collect()
