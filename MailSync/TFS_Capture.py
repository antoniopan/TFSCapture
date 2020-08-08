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


def modify_html(src_htm, dst_htm, src_xls, option):
    soup = bs4.BeautifulSoup(open(src_htm), features='html.parser')
    tables = soup.find_all('table')

    xls = xlrd.open_workbook(src_xls)

    # 读取User Requirement
    #fill_html_with_blank_row(tables[0], urSheet.nrows + 1)
    #sync_xls_html(urSheet, tables[0])

    # 读取P1 Task
    p1_task = xls.sheet_by_name("P1 Task List")
    fill_html_with_blank_row(tables[0], p1_task.nrows)
    fill_html_from_sheet(p1_task, tables[0])

    # 读取Task
    task_sheet = xls.sheet_by_name("Task Table")
    fill_html_with_blank_row(tables[1], task_sheet.nrows + 1)
    sync_xls_html(task_sheet, tables[1])

    # 读取本周Task
    task_this_week = xls.sheet_by_name("Task This Week")
    fill_html_with_blank_row(tables[2], task_this_week.nrows)
    fill_html_from_sheet(task_this_week, tables[2])

    # 读取CMTC UR
    ur_cmtc = xls.sheet_by_name("UR CMTC Table")
    fill_html_with_blank_row(tables[3], ur_cmtc.nrows + 1)
    sync_xls_html(ur_cmtc, tables[3])

    # 读取临床 UR
    ur_clinical = xls.sheet_by_name("UR Clinical Table")
    fill_html_with_blank_row(tables[4], ur_clinical.nrows + 1)
    sync_xls_html(ur_clinical, tables[4])

    # 读取本周UR
    ur_this_week = xls.sheet_by_name("UR This Week")
    fill_html_with_blank_row(tables[5], ur_this_week.nrows)
    fill_html_from_sheet(ur_this_week, tables[5])

    # 读取UnPlanned UR
    #fill_html_with_blank_row(tables[5], urUnPlannedSheet.nrows)
    #fill_html_from_sheet(urUnPlannedSheet, tables[5])

    # 读取Expired Task
    #fill_html_with_blank_row(tables[2], taskExpiredSheet.nrows)
    #fill_html_from_sheet(taskExpiredSheet, tables[2])

    # 读取UnReviewed Task
    #fill_html_with_blank_row(tables[7], taskUnReviewedSheet.nrows)
    #fill_html_from_sheet(taskUnReviewedSheet, tables[7])

    if option == 0:
        f = open(u'E:/Tracker/name.txt', 'r')
        s_name = f.read();
        f.close()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = s_name
        mail.Recipients.Add('liangliang.pan_HSW-GS')
        mail.Recipients.Add('HSW_DSA')
        mail.CC = 'jun.xiang_XR; qianqian.yu_XR; jinpeng.jiang_XR; ting.meng_XR; wanli.teng_XR; HSW_Manager'
        mail.Subject = 'DSA软件状态同步%s' % (time.strftime('%Y-%m-%d', time.localtime()))
        mail.BodyFormat = 2
        attachment = mail.Attachments.Add("E:/DSA_Software/image002.png")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ID001")
        attachment = mail.Attachments.Add("E:/DSA_Software/image004.png")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ID002")
        mail.HTMLBody = soup.prettify()
        mail.Send()
    elif option == 1:
        f = open(dst_htm, 'w', encoding='utf-8')
        f.write(soup.prettify())
        f.close()


def update_module_html(srcHtml, dstHtml, srcXls, option):
    soup = bs4.BeautifulSoup(open(srcHtml), features='html.parser')
    table = soup.table
    header = table.previous_sibling.previous_sibling
    xls_workbook = xlrd.open_workbook(srcXls)
    xls_sheets = xls_workbook.sheets()

    # Insert bland table and the title
    insertion_node = table.next_sibling.next_sibling

    i = 1
    while i < len(xls_sheets):
        insertion_node.insert_after(copy.copy(insertion_node))
        insertion_node.insert_after(copy.copy(table))
        insertion_node.insert_after(copy.copy(header))
        i += 1

    tables = soup.find_all('table')
    for i in range(0, len(tables)):
        table = tables[i]
        xls_sheet = xls_workbook.sheet_by_index(i)
        header = table.previous_sibling
        while header == '\n':
            header = header.previous_sibling
        header.span.string = "%s完成情况" % xls_sheet.name
        fill_html_with_blank_row(table, xls_sheet.nrows + 1)
        sync_xls_html(xls_sheet, table)

    if option == 0:
        f = open(u'E:/Tracker/name.txt', 'r')
        s_name = f.read();
        f.close()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = s_name
        mail.Recipients.Add('liangliang.pan_HSW-GS')
        mail.Recipients.Add('HSW_DSA')
        mail.CC = 'qianqian.yu_XR; baojian.wang_XR; ting.meng_XR; wanli.teng_XR; HSW_Manager'
        mail.Subject = 'DSA软件状态同步%s' % (time.strftime('%Y-%m-%d', time.localtime()))
        mail.BodyFormat = 2
        attachment = mail.Attachments.Add("E:/DSA_Software/image002.png")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ID001")
        attachment = mail.Attachments.Add("E:/DSA_Software/image004.png")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ID002")
        mail.HTMLBody = soup.prettify()
        mail.Send()
    elif option == 1:
        f = open(dstHtml, 'w', encoding='utf-8')
        f.write(soup.prettify())
        f.close()


def sync_xls_html(sheet, table):
    # 读取User Requirement
    rows = table.find_all('tr')
    if len(rows) - sheet.nrows != 2:
        return

    # Copy xls to html
    new = resolved = verified = 0
    for i in range(0, sheet.nrows):
        j = i + 1
        cols = rows[j].find_all('td')
        # 模块
        s = cols[0].find_all('span')
        s[0].string = sheet.cell(i, 0).value
        # 待解决
        s = cols[1].find_all('span')
        s[0].string = str(int(sheet.cell(i, 1).value))
        # 待验证
        s = cols[2].find_all('span')
        s[0].string = str(int(sheet.cell(i, 2).value))
        # 已验证
        s = cols[3].find_all('span')
        s[0].string = str(int(sheet.cell(i, 3).value))
        # 总和
        s = cols[4].find_all('span')
        s[0].string = str(int(sheet.cell(i, 4).value))
        # 解决率
        s = cols[5].find_all('span')
        s[0].string = '%.0f%%' % (sheet.cell(i, 5).value * 100)

        new = new + int(sheet.cell(i, 1).value)
        resolved = resolved + int(sheet.cell(i, 2).value)
        verified = verified + int(sheet.cell(i, 3).value)

    # 最后一行，统计
    cols = rows[len(rows) - 1].find_all('td')
    # 待解决
    s = cols[1].find_all('span')
    s[0].string = str(new)
    # 待验证
    s = cols[2].find_all('span')
    s[0].string = str(resolved)
    # 已验证
    s = cols[3].find_all('span')
    s[0].string = str(verified)
    # 总和
    s = cols[4].find_all('span')
    s[0].string = str(new + resolved + verified)
    # 解决率
    s = cols[5].find_all('span')
    s[0].string = '%.0f%%' % (float(resolved + verified) * 100 / (new + resolved + verified))


def fill_html_with_blank_row(table, nrows):
    if nrows < 3:
        n = nrows
        while n < 3:
            rows = table.find_all('tr')
            rows[1].decompose()
            n += 1
        return

    rows = table.find_all('tr')
    row1 = rows[1]
    row2 = rows[2]
    insertRow = rows[0]
    n = 3
    while n < nrows:
        if n % 2 == 1:
            newRow = copy.copy(row2)
        else:
            newRow = copy.copy(row1)
        insertRow.insert_after(newRow)
        n += 1


def fill_html_from_sheet(sheet, table):
    rows = table.find_all('tr')
    if len(rows) != (sheet.nrows + 1):
        print("row number mismatch.")
        return

    for i in range(0, sheet.nrows):
        cols = rows[i + 1].find_all('td')
        # ID
        s = cols[0].find('span')
        s.string = str(int(sheet.cell(i, 0).value))
        # Title
        s = cols[1].find('span')
        s.string = sheet.cell(i, 1).value
        # NodeName
        s = cols[2].find('span')
        s.string = sheet.cell(i, 3).value
        # AssignedTo
        s = cols[3].find('span')
        s.string = sheet.cell(i, 4).value
        # ExpectedSolvedDate
        if len(cols) > 4:
            s = cols[4].find('span')
            s.string = sheet.cell(i, 5).value
        if len(cols) > 5:
            s = cols[5].find('span')
            v = sheet.cell(i, 2).value
            if v != '':
                s.string = str(xlrd.xldate_as_datetime(v, 0).strftime('%y-%m-%d'))
            else:
                s.string = ''


if __name__ == '__main__':
    if len(sys.argv) == 2:
        f = open(sys.argv[1])
        lines = f.read().splitlines()
        day = datetime.date.today().isoweekday()
        # option 0 发邮件，option 1 更新html
        modify_html(lines[0], lines[1], lines[2], 0)
    else:
        f = open('./debug_file.txt', 'r')
        lines = f.read().splitlines()
        modify_html(lines[0], lines[1], lines[2], 1)
        update_module_html(lines[3], lines[4], lines[5], 1)
