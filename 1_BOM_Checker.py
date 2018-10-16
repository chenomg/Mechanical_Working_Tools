#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
# =============================================================================
#      FileName: 1_BOM_Checker.py
#          Desc: 检查BOM(第一张表格)表中文件名以及部件属性是否相匹配，第一列为标
识，为正整数时检查该行数据，第二列为文件名，结构：图号_名称，第三列为属性中图号，
第四列为属性中名称
#        Author: Jase Chen
#         Email: xxmm@live.cn
#      HomePage: https://jase.im/
#       Version: 0.0.1
#       License: GPLv2
#    LastChange: 2018-10-16 13:45:50
#       History:
# =============================================================================
'''

import xlrd
import re
import os

file_name = '_Sheet1.xlsx'
help_doc = """
用途：
    用于检查BOM(第一张表格)表中文件名以及部件属性是否相匹配
使用说明：
    BOM文件名：_Sheet1.xlsx
    BOM第一列为Flag，若为正整数则检查
    第二列为文件名，结构：图号_名称
    第三列为属性中图号
    第四列为属性中名称
"""

while True:
    try:
        os.system('cls')
        print(help_doc)
        book = xlrd.open_workbook(file_name)
        sheet = book.sheet_by_index(0)
        print('Total rows: ' + str(sheet.nrows))
        err_ID = 0
        err_name = 0
        reg = re.compile(r'^(.*)_(.*)$')
        for rx in range(sheet.nrows):
            try:
                # 若第一格为正整数才检查该行数据
                if int(sheet.cell(rx, 0).value) > 0:
                    # 文件名，格式：图号_名称
                    cell_1 = sheet.cell(rx, 1).value
                    # 验证图号是否一致
                    cell_2 = sheet.cell(rx, 2).value
                    # 验证名称是否一致
                    cell_3 = sheet.cell(rx, 3).value
                    if cell_1:
                        if reg.findall(cell_1)[0][0] != cell_2:
                            print(
                                'Row: {} '.format(rx + 1).ljust(12, '-') + ">",
                                'ID Wrong '.ljust(14, '-') + '>',
                                sheet.cell(rx, 1).value)
                            err_ID += 1
                        if reg.findall(cell_1)[0][1] != cell_3:
                            print(
                                'Row: {} '.format(rx + 1).ljust(12, '-') + ">",
                                'Name Wrong '.ljust(14, '-') + '>',
                                sheet.cell(rx, 1).value)
                            err_name += 1

            except Exception as e:
                pass

        if not err_ID and not err_name:
            print('All Checking Done.'.center(35, '+'))
        else:
            print("Wrong ID Count: {}".format(err_ID).center(35, '-'))
            print("Wrong Name Count: {}".format(err_name).center(35, '-'))
    except Exception as e:
        print(e)
    inp = input("Press 'Enter' to refresh, Press other key to exit --> ")
    if not inp:
        continue
    else:
        break
