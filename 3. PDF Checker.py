#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
# =============================================================================
#      FileName: PDF Checker.py
#          Desc: 用于检查PDF文件和BOM中的文件名是否能对应，显示出多余和缺少的PDF
#                第一列为flag，若为正整数则检查，若为-1则为对称件，不需要图纸
#                第二列为文件名
#                PDF文件和表格需要放在和本程序同一文件夹
#        Author: Jase Chen
#         Email: xxmm@live.cn
#      HomePage: https://jase.im/
#       Version: 0.0.1
#       License: GPLv2
#    LastChange: 2018-09-13 17:28:05
#       History:
# =============================================================================
'''
import os
import re
import xlrd
import time

path = os.getcwd()
xl_file = '_Sheet1.xlsx'


def get_PDF_files(path):
    contents = os.listdir(path)
    files = []
    for f in contents:
        if not os.path.isdir(f):
            if re.findall(r'^(.*)\.(\w+)$', f):
                name_ext = re.findall(r'^(.*)\.(\w+)$', f)[0]
                if name_ext[1].lower() == 'pdf':
                    files.append(name_ext[0])
    return files


def get_checked_files_inXLS(xl_file):
    try:
        book = xlrd.open_workbook(xl_file)
        sheet = book.sheet_by_index(0)
        values = set()
        symmetry_files = []
        for row in range(sheet.nrows):
            try:
                if int(sheet.cell_value(row, 0)) > 0:
                    values.add(sheet.cell_value(row, 1))
                if int(sheet.cell_value(row, 0)) == -1:
                    symmetry_files.append(sheet.cell_value(row, 1))
            except:
                pass
        return values, symmetry_files
    except Exception as e:
        return (None, None)
        print(e)


def main():
    while True:

        start_time = time.time()
        os.system('cls')
        files = get_PDF_files(path)
        values, symmetry_files = get_checked_files_inXLS(xl_file)
        files_useless = []
        files_lost = []
        if files and values:
            for f in files:
                if f in values:
                    continue
                else:
                    files_useless.append(f)
            for v in values:
                if v in files:
                    continue
                else:
                    files_lost.append(v)
        print('\nPDF图纸文件和BOM对比结果如下↓')
        print('\n' +
              'Files useless: {}'.format(len(files_useless)).center(50, '-'))
        if not files_useless:
            print('未发现多余的文件')
        else:
            for i in files_useless:
                print(i)
            print('-' * 50)
        print('\n' + 'Files lost: {}'.format(len(files_lost)).center(50, '-'))
        if not files_lost:
            print('未发现缺少的文件')
        else:
            for i in files_lost:
                print(i)
            print('-' * 50)
        print('\n' +
              'Symmetry files: {}'.format(len(symmetry_files)).center(50, '-'))
        if not symmetry_files:
            print('未找到使用对称图纸的文件')
        else:
            for i in symmetry_files:
                print(i)
            print('-' * 50)
        end_time = time.time()
        print('\n过程耗时：{}s'.format(round(end_time - start_time, 2)))
        inp = input('\nPress "Enter" to refresh, press other key to Exit!')
        if inp:
            break


if __name__ == "__main__":
    main()
