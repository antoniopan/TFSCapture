# -*- coding: utf-8 -*-
"""
Created on Tue Apr  7 06:56:55 2020

@author: liangliang.pan
"""

import bs4
import xlrd
import copy
import time
import win32com.client as win32


def ModifyHtml(srcFile, dstFile, xlsFile):
    soup = bs4.BeautifulSoup(open(srcFile), features='html.parser')
    tables = soup.find_all('table')

    xls = xlrd.open_workbook(xlsFile)
    urSheet = xls.sheet_by_index(0)
    taskSheet = xls.sheet_by_index(1)
    urExpiredSheet = xls.sheet_by_index(2)
    urUnPlannedSheet = xls.sheet_by_index(3)
    taskExpiredSheet = xls.sheet_by_index(4)
    taskUnReviewedSheet = xls.sheet_by_index(5)
    urThisWeek = xls.sheet_by_index(6)
    taskThisWeek = xls.sheet_by_index(7)

    # 读取User Requirement
    FillHtmlWithBlankRow(tables[0], urSheet.nrows + 1)
    SyncXlsHtml(urSheet, tables[0])

    # 读取Task
    FillHtmlWithBlankRow(tables[1], taskSheet.nrows + 1)
    SyncXlsHtml(taskSheet, tables[1])

    # 读取本周UR
    FillHtmlWithBlankRow(tables[2], urThisWeek.nrows)
    FillHtmlFromSheet(urThisWeek, tables[2])

    # 读取本周Task
    FillHtmlWithBlankRow(tables[3], taskThisWeek.nrows)
    FillHtmlFromSheet(taskThisWeek, tables[3])

    # 读取Expired UR
    FillHtmlWithBlankRow(tables[4], urExpiredSheet.nrows)
    FillHtmlFromSheet(urExpiredSheet, tables[4])

    # 读取UnPlanned UR
    FillHtmlWithBlankRow(tables[5], urUnPlannedSheet.nrows)
    FillHtmlFromSheet(urUnPlannedSheet, tables[5])

    # 读取Expired Task
    FillHtmlWithBlankRow(tables[6], taskExpiredSheet.nrows)
    FillHtmlFromSheet(taskExpiredSheet, tables[6])

    # 读取UnReviewed Task
    FillHtmlWithBlankRow(tables[7], taskUnReviewedSheet.nrows)
    FillHtmlFromSheet(taskUnReviewedSheet, tables[7])


    f = open(dstFile, 'w', encoding='utf-8')
    f.write(soup.prettify())
    f.close()
    '''
    f = open(u'E:/Tracker/name.txt', 'r')
    sName = f.read();
    f.close()

    # sys.setdefaultencodeing('utf-8')
    # warnings.filterwarnings('ignore')
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    #mail.To = sName
    mail.Recipients.Add('liangliang.pan_HSW-GS')
    #mail.Recipients.Add('HSW_DSA')
    #mail.CC = 'qianqian.yu_XR; baojian.wang_XR; ting.meng_XR; wanli.teng_XR; HSW_Manager'
    mail.Subject = 'DSA软件状态同步%s' % (time.strftime('%Y-%m-%d', time.localtime()))
    mail.BodyFormat = 2
    attachment = mail.Attachments.Add("E:/DSA_Software/image002.png")
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ID001")
    attachment = mail.Attachments.Add("E:/DSA_Software/image004.png")
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ID002")
    mail.HTMLBody = soup.prettify()
    mail.Send()
    '''

def SyncXlsHtml(sheet, table):
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


def FillHtmlWithBlankRow(table, nrows):
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


def FillHtmlFromSheet(sheet, table):
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
            v = sheet.cell(i, 2).value
            if v != '':
                s.string = str(xlrd.xldate_as_datetime(v, 0).strftime('%y-%m-%d'))
            else:
                s.string = ''


ModifyHtml(u'E:/Tracker/DSA_Software_Daily_Tracker.htm',
           u'E:/Tracker/DSA_Software_Daily_Tracker_v1.htm',
           u'E:/Tracker/temp.xlsx')
