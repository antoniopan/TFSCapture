import bs4
import xlrd
import sys
import time
import datetime
import win32com.client as win32
import modify_table
import os


def modify_html(src_htm, dst_htm, src_xls, name_file, option):
    soup = bs4.BeautifulSoup(open(src_htm), features='html.parser')
    tables = soup.find_all('table')

    xls = xlrd.open_workbook(src_xls)

    # 读取自测Task
    task_sheet = xls.sheet_by_name("Improvement Task Expired")
    modify_table.fill_html_with_blank_row(tables[0], task_sheet.nrows)
    modify_table.fill_html_from_sheet(task_sheet, tables[0])

    # 读取未评审H3 Improvement Task
    task_sheet = xls.sheet_by_name("Improvement Task Unreviewed")
    modify_table.fill_html_with_blank_row(tables[1], task_sheet.nrows)
    modify_table.fill_html_from_sheet(task_sheet, tables[1])

    # 读取未评审H3 Improvement Task
    task_sheet = xls.sheet_by_name("Improvement Task Not Planned")
    modify_table.fill_html_with_blank_row(tables[2], task_sheet.nrows)
    modify_table.fill_html_from_sheet(task_sheet, tables[2])

    # 读取本周解决 Task
    task_sheet = xls.sheet_by_name("Task Change This Week")
    modify_table.fill_html_with_blank_row(tables[3], task_sheet.nrows)
    modify_table.fill_html_from_sheet_create_resolve(task_sheet, tables[3])

    if os.path.exists(dst_htm):
        os.remove(dst_htm)

    if option == 0:
        f = open(name_file, 'r')
        s_name = f.read()
        f.close()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = s_name
        mail.Recipients.Add('liangliang.pan_HSW-GS')
        mail.CC = 'HSW_GS_IPA_AP-APP2; HSW_GS_IPA_AP_APPCOM'
        mail.Subject = 'UIDealB69SP4H4软件状态同步%s' % (time.strftime('%Y-%m-%d', time.localtime()))
        mail.HTMLBody = soup.prettify()
        mail.Send()
    elif option == 1:
        f = open(dst_htm, 'w', encoding='utf-8')
        f.write(soup.prettify())
        f.close()


if __name__ == '__main__':
    if len(sys.argv) == 2:
        f = open(sys.argv[1])
        lines = f.read().splitlines()
        day = datetime.date.today().isoweekday()
        modify_html(lines[0], lines[1], lines[2], lines[3], 0)
    else:
        f = open('./uideal_b69sp4.txt', 'r')
        lines = f.read().splitlines()
        modify_html(lines[0], lines[1], lines[2], lines[3], 1)