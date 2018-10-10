#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
# =============================================================================
#      FileName: BOM Checker - project.py
#          Desc: 检查BOM和由SolidWorks导出的清单进行校对，核对BOM中项目有无缺失，
                 以及数量是否正. 确第一列为标识，为正整数时检查该行数据，
                 第二列为图号，第三列为数量
#        Author: Jase Chen
#         Email: xxmm@live.cn
#      HomePage: https://jase.im/
#       Version: 0.0.1
#       License: GPLv2
#    LastChange: 2018-10-10 14:40:03
#       History:
# =============================================================================
'''

import xlrd
import re
import os
import time

file_name_project = 'FBCA.00.20D_人工供包机(900LD)V2.0.xlsx'
file_name_checker = 'FBCA.00.20D_人工供包机(900LD)V2.0 - checker.xlsx'
while True:
    try:
        start_time = time.time()
        os.system('cls')
        book_p = xlrd.open_workbook(file_name_project)
        sheet_p = book_p.sheet_by_index(0)
        book_c = xlrd.open_workbook(file_name_checker)
        sheet_c = book_c.sheet_by_index(0)
        # BOM数据
        data_p = {}
        # checker数据
        data_c = {}
        item_lost = []
        item_useless = []
        quantity_wrong = []
        for rx in range(sheet_p.nrows):
            try:
                # 若第一格为正整数才检查该行数据
                if int(sheet_p.cell(rx, 0).value) > 0:
                    # 图号
                    cell_1 = sheet_p.cell(rx, 1).value
                    if cell_1:
                        # 单套数量
                        cell_2 = sheet_p.cell(rx, 2).value
                        data_p[cell_1] = cell_2

            except Exception as e:
                pass

        for rx in range(sheet_c.nrows):
            try:
                # 若第一格为正整数才检查该行数据
                if int(sheet_c.cell(rx, 0).value) > 0:
                    # 图号
                    cell_1 = sheet_c.cell(rx, 1).value
                    if cell_1:
                        # 单套数量
                        cell_2 = sheet_c.cell(rx, 2).value
                        data_c[cell_1] = cell_2

            except Exception as e:
                pass

        for item in data_c:
            if not data_p.get(item):
                item_lost.append(item)
            else:
                if not data_c.get(item) == data_p.get(item):
                    quantity_wrong.append(item)
        for item in data_p:
            if data_p.get(item):
                if not data_c.get(item):
                    item_useless.append(item)
        print('缺少的项：')
        for i in item_lost:
            print(i)
        print('数量错误的项：')
        for i in quantity_wrong:
            print(i)
        print('多余的项：')
        for i in item_useless:
            print(i)
        # if not err_ID and not err_name:
        # print('All Checking Done.'.center(35, '+'))
        # else:
        # print("Wrong ID Count: {}".format(err_ID).center(35, '-'))
        # print("Wrong Name Count: {}".format(err_name).center(35, '-'))
    except Exception as e:
        print(e)
    end_time = time.time()
    print('过程耗时：{}s'.format(round((end_time - start_time), 2)))
    inp = input("Press 'Enter' to refresh, Press other key to exit --> ")
    if not inp:
        continue
    else:
        break
