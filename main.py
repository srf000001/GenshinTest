import os
import re
import sys

import xlrd
import xlwt
from xlutils.copy import copy

# pip install PyInstaller -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com

Judgeline = [r'^\S{1,4}$', r'^[1-9]\d{0,4}$', r'^\d{0,3}[.]\d{0,2}|\d{0,3}$', r'^[1-9]\d{0,2}[.]\d{0,2}|\d{0,3}$',
             r'^[1-9]\d{0,2}[.]\d{0,2}|\d{0,3}$',
             r'^[1-9]\d{0,2}[.]\d{0,2}|\d{0,3}$', r'^[1-9]\d{0,2}[.]\d{0,2}|\d{0,3}$',
             r'^[1-9]\d{0,3}[.]\d{0,2}|\d{0,4}$', r'^[1-9]\d{0,3}[.]\d{0,2}|\d{0,5}$']
Trueline = ['名称为1到4个字符', '主属性必须为数字1-9999', '元素伤害加成必须为0-999']


# AllLine前两位是行数和列数，第三位开始为属性
def ExcelOpen(excelname='Genshin.xls', AllLine=[]):
    rows = AllLine[0]
    cols = AllLine[1]
    f = xlrd.open_workbook(excelname)
    wb = copy(f)
    ws = wb.get_sheet(0)
    line = []
    for i in range(0, cols - 1):
        str = '请输入' + AllLine[i + 2] + ': '
        k = True
        while k:
            a = input(str)
            if a == '000':
                print('你中断了输入')
                return
            m1 = re.match(Judgeline[i], a)
            if m1:
                k = False
            else:
                print('输入格式错误，请重新输入！')
        line.append(a)
    line.append(DamageCaculate(line))
    for x in range(0, len(line)):
        ws.write(rows, x, line[x])
    wb.save(excelname)


def DamageCaculate(line):
    a = float(line[1])
    b = float(100 + float(line[2])) / 100
    c = float(line[3]) / 100
    d = float(float(line[4])) / 100
    sum = a * b * (c * d + 1)
    return str(sum)


# import xlwt
#
# workbook = xlwt.Workbook(encoding='utf-8')  # 新建工作簿
# sheet1 = workbook.add_sheet("测试表格")  # 新建sheet
# workbook.save(r'D:\PycharmProjects\test.xlsx')  # 保存


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def lineGet(excelname='Genshin.xls', wantprint='0'):
    try:
        f = xlrd.open_workbook(excelname)
    except FileNotFoundError:
        print('无法打开指定的文件!')
        xls = xlwt.Workbook()
        sht1 = xls.add_sheet('sheet1')
        xls.save(excelname)
    except LookupError:
        print('指定了未知的编码!')
        return
    except UnicodeDecodeError:
        print('读取文件时解码错误!')
        return
    finally:
        print('文件打开成功!')
    # 打开第n个sheet表格
    table = f.sheets()[0]
    # print('行数', table.nrows)
    # print('列数', table.ncols)
    line = [table.nrows, table.ncols]
    line += table.row_values(0)
    line.pop(len(line) - 1)
    if wantprint != '0':
        for x in range(0, table.nrows):
            line1 = table.row_values(x)
            if x == 0:
                print(line1)
            else:
                for y in line1:
                    print('%5s' % y, end='\t')
                print('\n')

    # for x in line:
    #     print('属性：', x)
    return line


def Open():
    str = get_resource_path('Genshin.xls')
    while True:
        print('1：查看当前表格\n2：查看表格后输入新数据\n3：退出程序')
        a = int(input('请输入代码：'))
        if a == 1:
            lineGet(excelname=str, wantprint='1')
        elif a == 2:
            line = lineGet(excelname=str, wantprint='1')
            ExcelOpen(excelname=str, AllLine=line)
            lineGet(excelname=str, wantprint='1')
        elif a == 3:
            return


def main():
    Open()
    # os.system('pause')

    # ExcelOpen(AllLine=line)


if __name__ == '__main__':
    main()
