# coding: utf-8

import csv
import datetime
import os
import re
import argparse

import xlrd


def return_sheetnames(xlsxfile):
    sheetlist = []
    book = xlrd.open_workbook(xlsxfile)
    for sheet in book.sheets():
        sheet_nm = sheet.name
        sheetlist.append(sheet_nm)
    return sheetlist




def sheet2tsv(xlsxfile, sheetname, outputfile, dt=False, sep=',', skip=False, encoding='Shift-JIS'):

    book = xlrd.open_workbook(xlsxfile)
    sheet = book.sheet_by_name(sheetname)

    if os.path.exists(outputfile):
        os.remove(outputfile)

    with open(outputfile, 'a', encoding=encoding, newline="\n") as fp:
        writer = csv.writer(fp, delimiter=sep)

        for row in range(sheet.nrows):

            if skip:
                if row == 0:
                    continue

            li = []
            for col in range(sheet.ncols):
                cell = sheet.cell(row, col)

                if cell.ctype == xlrd.XL_CELL_ERROR:
                    val = ""
                    continue

                if cell.ctype == xlrd.XL_CELL_NUMBER:
                    val = cell.value

                    if val.is_integer():
                        val = int(val)

                elif cell.ctype == xlrd.XL_CELL_DATE:
                    d = get_dt_from_serial(cell.value)
                    if dt:
                        val = d.strftime('%Y/%m/%d')
                    else:
                        val = d.strftime('%Y/%m/%d %H:%M:%S')

                else:  # その他
                    val = cell.value

                li.append(val)
            writer.writerow(li)

    if os.path.getsize(outputfile) == 0:
        os.remove(outputfile)



def get_dt_from_serial(serial):
    if serial == 0:
        return None

    base_date = datetime.datetime(1899, 12, 30)
    d, t = re.search(r'(\d+)(\.\d+)', str(serial)).groups()
    return base_date + datetime.timedelta(days=int(d)) \
        + datetime.timedelta(seconds=float(t) * 86400)