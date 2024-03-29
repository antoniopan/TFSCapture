import copy
import xlrd
import math


def fill_html_with_blank_row(table, number_rows):
    if number_rows < 3:
        n = number_rows
        while n < 3:
            rows = table.find_all('tr')
            rows[1].decompose()
            n += 1
        return

    rows = table.find_all('tr')
    row1 = rows[1]
    row2 = rows[2]
    insert_row = rows[0]
    n = 3
    while n < number_rows:
        if n % 2 == 1:
            new_row = copy.copy(row2)
        else:
            new_row = copy.copy(row1)
        insert_row.insert_after(new_row)
        n += 1


def sync_xls_html(sheet, table):
    # 读取User Requirement
    rows = table.find_all('tr')
    if len(rows) - sheet.nrows != 1:
        return

    # Copy xls to html
    new = resolved = verified = 0
    for i in range(0, sheet.nrows):
        j = i + 1
        cols = rows[j].find_all('td')
        # 模块
        s = cols[0].find_all('span')
        s[0].string = sheet.cell(i, 0).value
        for k in range(1, len(cols) - 3):
            s = cols[k].find_all('span')
            s[0].string = str(int(sheet.cell(i, k).value))

        k = sheet.ncols - 2
        # 验证率
        s = cols[k + 1].find_all('span')
        d_percentage = sheet.cell(i, k + 1).value
        s[0].string = '%.0f%%' % (d_percentage * 100)
        # 解决率
        s = cols[k].find_all('span')
        d_percentage = sheet.cell(i, k).value
        s[0].string = '%.0f%%' % (d_percentage * 100)
        # 待解决个数
        s = cols[k + 2].find_all('span')
        n_bugs = math.ceil((0.9 - d_percentage) * sheet.cell(i, k - 1).value)
        s[0].string = '%d' % n_bugs

        new = new + int(sheet.cell(i, 1).value)
        resolved = resolved + int(sheet.cell(i, 2).value)
        verified = verified + int(sheet.cell(i, 3).value)


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
        # Priority
        s = cols[2].find('span')
        if sheet.cell(i, 2).value != '':
            s.string = str(int(sheet.cell(i, 2).value))
        # NodeName
        s = cols[3].find('span')
        s.string = sheet.cell(i, 3).value
        # AssignedTo
        s = cols[4].find('span')
        s.string = sheet.cell(i, 4).value
        # ExpectedSolvedDate
        j = len(cols)
        s = cols[j - 1].find('span')
        if sheet.ncols <= 5:
            s.string = ''
        else:
            v = sheet.cell(i, 5).value
            if v != '':
                s.string = str(xlrd.xldate_as_datetime(v, 0).strftime('%y-%m-%d'))
            else:
                s.string = ''


def fill_html_from_sheet_create_resolve(sheet, table):
    rows = table.find_all('tr')
    if len(rows) != (sheet.nrows + 1):
        print("row number mismatch.")
        return

    for i in range(0, sheet.nrows):
        cols = rows[i + 1].find_all('td')
        # Node Name
        s = cols[0].find('span')
        s.string = sheet.cell(i, 0).value
        # Create Number
        s = cols[1].find('span')
        s.string = str(int(sheet.cell(i, 1).value))
        # Resolve Number
        s = cols[2].find('span')
        s.string = str(int(sheet.cell(i, 2).value))
