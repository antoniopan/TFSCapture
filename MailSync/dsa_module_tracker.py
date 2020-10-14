import bs4
import xlrd
import copy
import sys
import time
import datetime
import win32com.client as win32
import modify_table
import utils


def update_module_html(srcHtml, dstHtml, srcXls, img_dir, option):
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
        modify_table.fill_html_with_blank_row(table, xls_sheet.nrows)
        modify_table.sync_xls_html(xls_sheet, table)

    if option == 0:
        images = ["%s/ProjPlan.png" % img_dir, "%s/image004.png" % img_dir]
        print(images)
        # receiver = ''
        receiver = 'HSW_DSA'
        # CC = ''
        CC = 'jun.xiang_XR; qianqian.yu_XR; jinpeng.jiang_XR; ting.meng_XR; wanli.teng_XR; HSW_Manager'
        utils.send_mail('', 'DSA软件状态同步', soup.prettify(), receiver, CC, images)

    elif option == 1:
        f = open(dstHtml, 'w', encoding='utf-8')
        f.write(soup.prettify())
        f.close()


if __name__ == '__main__':
    if len(sys.argv) == 2:  # 参数齐全，发邮件
        f = open(sys.argv[1])
        lines = f.read().splitlines()
        day = datetime.date.today().isoweekday()
        # option 0 发邮件，option 1 更新html
        update_module_html(lines[3], lines[4], lines[5], lines[6], 0)
    else:  # 本地运行，生成html文件
        f = open('./debug_file.txt', 'r')
        lines = f.read().splitlines()
        update_module_html(lines[3], lines[4], lines[5], lines[6], 1)
