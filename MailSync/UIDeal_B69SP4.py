import bs4
import xlrd
import sys
import time
import datetime
import win32com.client as win32
import modify_table


def modify_html(src_htm, dst_htm, src_xls, name_file, option):
    soup = bs4.BeautifulSoup(open(src_htm), features='html.parser')
    tables = soup.find_all('table')

    xls = xlrd.open_workbook(src_xls)

    # 读取过期Improvement Task
    task_sheet = xls.sheet_by_name("Improvement Task Expired")
    modify_table.fill_html_with_blank_row(tables[0], task_sheet.nrows)
    modify_table.fill_html_from_sheet(task_sheet, tables[0])

    # 读取未评审Improvement Task
    task_sheet = xls.sheet_by_name("Improvement Task Unreviewed")
    modify_table.fill_html_with_blank_row(tables[1], task_sheet.nrows)
    modify_table.fill_html_from_sheet(task_sheet, tables[1])

    # 读取过期Design Task
    task_sheet = xls.sheet_by_name("Designed Task Expired")
    modify_table.fill_html_with_blank_row(tables[2], task_sheet.nrows)
    modify_table.fill_html_from_sheet(task_sheet, tables[2])

    if option == 0:
        f = open(name_file, 'r')
        s_name = f.read();
        f.close()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = s_name
        mail.Recipients.Add('liangliang.pan_HSW-GS')
        mail.CC.Add('HSW_GS_IPA_AP')
        mail.Subject = 'UIDealB69SP4H2软件状态同步%s' % (time.strftime('%Y-%m-%d', time.localtime()))
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