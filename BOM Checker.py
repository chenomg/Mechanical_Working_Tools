#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
# =============================================================================
#      FileName: 0check.py
#          Desc: 检查BOM(第一张表格)表中文件名以及部件属性是否相匹配，第一列为标
识，为正整数时检查该行数据，第二列为文件名，结构：图号_名称，第三列为属性中图号，
第四列为属性中名称
#        Author: Jase Chen
#         Email: xxmm@live.cn
#      HomePage: https://jase.im/
#       Version: 0.0.1
#       License: GPLv2
#    LastChange: 2018-06-06 13:33:40
#       History:
# =============================================================================
'''

import xlrd
import re

file_name = 'FBCC.00.10J_人工供包机(900RJ)V2.0.xlsx'
try:
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
                    if reg.findall(cell_1)[0][0] == cell_2:
                        pass
                    else:
                        print('ID wrong in row: ' + str(rx + 1),
                              sheet.cell(rx, 1).value)
                        err_ID += 1
                if cell_1:
                    if reg.findall(cell_1)[0][1] == cell_3:
                        pass
                        # print('OK')
                    else:
                        print('Name wrong in row: ' + str(rx + 1),
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
