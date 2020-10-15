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

    # 读取User Requirement
    #fill_html_with_blank_row(tables[0], urSheet.nrows + 1)
    #sync_xls_html(urSheet, tables[0])

    # 读取P1 Task
    #p1_task = xls.sheet_by_name("P1 Task List")
    #fill_html_with_blank_row(tables[0], p1_task.nrows)
    #fill_html_from_sheet(p1_task, tables[0])

    # 读取Task
    task_sheet = xls.sheet_by_name("Task Table")
    modify_table.fill_html_with_blank_row(tables[0], task_sheet.nrows)
    modify_table.sync_xls_html(task_sheet, tables[0])

    # 读取本周Task
    task_this_week = xls.sheet_by_name("Task This Week")
    modify_table.fill_html_with_blank_row(tables[1], task_this_week.nrows)
    modify_table.fill_html_from_sheet(task_this_week, tables[1])

    # 读取CMTC UR
    ur_all = xls.sheet_by_name("UR Table")
    modify_table.fill_html_with_blank_row(tables[2], ur_all.nrows)
    modify_table.sync_xls_html(ur_all, tables[2])

    # 读取临床 UR
    #ur_clinical = xls.sheet_by_name("UR Clinical Table")
    #fill_html_with_blank_row(tables[4], ur_clinical.nrows)
    #sync_xls_html(ur_clinical, tables[4])

    # 读取本周UR
    ur_this_week = xls.sheet_by_name("UR This Week")
    modify_table.fill_html_with_blank_row(tables[3], ur_this_week.nrows)
    modify_table.fill_html_from_sheet(ur_this_week, tables[3])

    # 读取Expired UR
    urExpired = xls.sheet_by_name("Expired UR")
    modify_table.fill_html_with_blank_row(tables[4], urExpired.nrows)
    modify_table.fill_html_from_sheet(urExpired, tables[4])

    # 读取Expired Task
    #fill_html_with_blank_row(tables[2], taskExpiredSheet.nrows)
    #fill_html_from_sheet(taskExpiredSheet, tables[2])

    # 读取UnReviewed Task
    #fill_html_with_blank_row(tables[7], taskUnReviewedSheet.nrows)
    #fill_html_from_sheet(taskUnReviewedSheet, tables[7])

    if option == 0:
        images = ["%s/ProjPlan.png" % img_dir, "%s/image004.png" % img_dir]
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
