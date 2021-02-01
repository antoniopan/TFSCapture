# -*- coding: utf-8 -*-
"""
Created on Tue Apr  7 06:56:55 2020

@author: liangliang.pan
"""

import bs4
import xlrd
import copy
import sys
import time
import datetime
import win32com.client as win32
import modify_table
import utils

def modify_html(src_htm, dst_htm, src_xls, name_file, img_dir, option):
    soup = bs4.BeautifulSoup(open(src_htm), features='html.parser')
    tables = soup.find_all('table')

    xls = xlrd.open_workbook(src_xls)

    # 读取Task
    task_sheet = xls.sheet_by_name("Task Table")
    modify_table.fill_html_with_blank_row(tables[0], task_sheet.nrows)
    modify_table.sync_xls_html(task_sheet, tables[0])

    # 读取Expired Task
    ur_this_week = xls.sheet_by_name("Task Expired")
    modify_table.fill_html_with_blank_row(tables[1], ur_this_week.nrows)
    modify_table.fill_html_from_sheet(ur_this_week, tables[1])

    # 读取Unreviewed UR
    urExpired = xls.sheet_by_name("Task Unreviewed")
    modify_table.fill_html_with_blank_row(tables[2], urExpired.nrows)
    modify_table.fill_html_from_sheet(urExpired, tables[2])

    # 读取本周Task变化
    task_this_week = xls.sheet_by_name("Task Change This Week")
    modify_table.fill_html_with_blank_row(tables[3], task_this_week.nrows)
    modify_table.fill_html_from_sheet_create_resolve(task_this_week, tables[3])

    if option == 0:
        images = ["%s/Bug评审流程.png" % img_dir]
        # receiver = ''
        receiver = 'HSW_DSA'
        # CC = ''
        CC = 'jun.xiang_XR; qianqian.yu_XR; jinpeng.jiang_XR; ting.meng_XR; wanli.teng_XR; HSW_Manager'
        utils.send_mail(name_file, 'DSA软件状态同步', soup.prettify(), receiver, CC, images)
    elif option == 1:
        f_name = open(dst_htm, 'w', encoding='utf-8')
        f_name.write(soup.prettify())
        f_name.close()


if __name__ == '__main__':
    if len(sys.argv) == 2: # 参数齐全，发邮件
        f = open(sys.argv[1])
        lines = f.read().splitlines()
        day = datetime.date.today().isoweekday()
        # option 0 发邮件，option 1 更新html
        modify_html(lines[0], lines[1], lines[2], lines[7], lines[6], 0)
    else: # 本地运行，生成html文件
        f = open('./debug_file.txt', 'r')
        lines = f.read().splitlines()
        modify_html(lines[0], lines[1], lines[2], lines[6], lines[7], 1)
