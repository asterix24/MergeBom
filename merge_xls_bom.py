# MergeBom is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
#
# Copyright 2012 Daniele Basile <asterix24@gmail.com>
#
#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import sys, os
import re

import xlrd
import xlsxwriter
import datetime

def fillRowCenter(row, s):
    r = row[:(len(row)/2) - (len(s)/2)] + s + row[(len(row)/2) + (len(s)/2):]
    if len(r) < len(row):
        r = r + row[-1] * (len(row) - len(r))
    if len(r) > len(row):
        r = r[:-1 * (len(r) - len(row))]
    return r

def fillTableRow(row, col1, col2, col3):
    col1 = "%s" % col1
    col2 = "%s" % col2
    col3 = "%s" % col3

    if len(col1) > 10:
        col1 = col1[:10]
    s = col1 + row[10 - len(col1):]

    WCOL2=40
    if len(col2) > WCOL2:
        col2 = col2[:WCOL2]
        print "trim col2"
    s = s[:(WCOL2-len(col2))] + col2 + s[WCOL2:]

    if len(col3) > (len(row) - WCOL2):
        col3 = col3[:len(row) - WCOL2]
        print "trim col3"

    s = s[: len(row) - len(col3)] + col3

    return s



valid_keys = [u'quantity', u'designator', u'comment', u'footprint', u'description']
connectors_designators = ['J', 'X', 'P', 'SIM']

def parse_data(bom_file_list):
    CATEGORY = 0
    TOT_COUNT = 1
    XLS_DSG = 0
    XLS_COM = 1
    XLS_FOT = 2
    XLS_DSC = 3

    FILES = len(bom_file_list)
    FIELDS = len(valid_keys)
    STAT_FIELDS = len([CATEGORY, TOT_COUNT])

    table_dict = {}
    header = []

    for index_file,i in enumerate(bom_file_list):
        wb, data = read_xls(i)

        header = filter(lambda x: x[0].lower() in valid_keys, data)
        header = header[0] if len(header) else []

        d = {}
        for n,i in enumerate(header):
            d[i.lower()] = n

        QUANTITY=d['quantity']
        DESIGNATOR=d['designator']
        COMMENT=d['comment']
        FOOTPRINT=d['footprint']
        DESCRIPTION=d['description']


        data = filter(lambda x: x[0].lower() not in valid_keys, data)
        for j in data:
            FIX_LEN = abs((STAT_FIELDS + FILES + FIELDS) - len(j))
            if filter(lambda x: x, j):
                key = ""
                for i in connectors_designators:
                    if i in j[DESIGNATOR].strip().upper():
                        key = j[FOOTPRINT] + j[DESCRIPTION]
                        print "!! Key Merged: > ",key
                    else:
                        key = j[COMMENT] + j[FOOTPRINT] + j[DESCRIPTION]

                if key == '':
                    print "~" * 80
                    print "NULL KEY ERROR"
                    print j
                    print "~" * 80
                    continue

                # Calcolo il tipo di componente.

                if re.findall("\S,[\S]+", j[DESIGNATOR]):
                    j[DESIGNATOR] = j[DESIGNATOR].replace(",", ", ")

                c = re.search('^[a-zA-Z_]{1,3}', j[DESIGNATOR])
                group_key = ''
                if c is not None:
                    group_key = c.group().upper()
                    # Buttons and spacer
                    if group_key in ['B', 'BT', 'SCR', 'SPA', 'BAT','SW']:
                        group_key = 'S'
                    # Fuses
                    if group_key in ['G']:
                        group_key = 'F'
                    # Tranformer
                    if group_key in ['T' ]:
                        group_key = 'TR'
                    # Resistors, array, etc.
                    if group_key in ['RN', 'R_G']:
                        group_key = 'R'
                    # Connectors
                    if group_key in ['X', 'P', 'SIM']:
                        group_key = 'J'
                    # Discarted ref
                    if group_key in ['TP']:
                        print "WARNING WE SKIP THIS KEY"
                        print 'key [%s]' % group_key
                        print '.........'
                        continue
                else:
                    print "GROUP key not FOUND!"
                    print j
                    sys.exit(1)

                raw_row = [''] * (STAT_FIELDS + FILES + FIELDS + FIX_LEN)
                raw_row[CATEGORY] = group_key

                if table_dict.has_key(key):
                    raw_row = table_dict[key]

                    raw_row[TOT_COUNT] = int(j[QUANTITY]) + int(table_dict[key][TOT_COUNT])
                    raw_row[STAT_FIELDS + index_file]      = int(j[QUANTITY])
                    raw_row[STAT_FIELDS + FILES + XLS_DSG] = ", ".join([j[DESIGNATOR], table_dict[key][STAT_FIELDS + FILES + XLS_DSG]])
                    raw_row[STAT_FIELDS + FILES + XLS_COM] = j[COMMENT]
                    raw_row[STAT_FIELDS + FILES + XLS_FOT] = j[FOOTPRINT]
                    raw_row[STAT_FIELDS + FILES + XLS_DSC] = j[DESCRIPTION]

                else:
                    raw_row[TOT_COUNT] = int(j[QUANTITY])
                    raw_row[STAT_FIELDS + index_file]      = int(j[QUANTITY])
                    raw_row[STAT_FIELDS + FILES + XLS_DSG] = j[DESIGNATOR]
                    raw_row[STAT_FIELDS + FILES + XLS_COM] = j[COMMENT]
                    raw_row[STAT_FIELDS + FILES + XLS_FOT] = j[FOOTPRINT]
                    raw_row[STAT_FIELDS + FILES + XLS_DSC] = j[DESCRIPTION]

                    table_dict[key] = raw_row

            file_list = map(os.path.basename, bom_file_list)
            table_dict['header'] = ["Tot.Qty"] + file_list + header

    header = table_dict['header']
    del table_dict['header']
    return header, table_dict.values()



ORDER_PATTERN = ['J', 'S', 'F','R','C','D','DZ','L', 'Q','TR','Y', 'U']
ORDER_PATTERN_NAMES = {
    'J':['* J Connectors *'],
    'S':['* S Mechanical parts and buttons *'],
    'F':['* F Fuses *'],
    'R':['* R Resistors *'],
    'C':['* C Capacitors *'],
    'D':['* D Diodes *'],
    'DZ':['* DZ Zener, Schottky, Transil *'],
    'L': ['* L Inductors, chokes *'],
    'Q': ['* Q Transistors *'],
    'TR':['* TR Transformers *'],
    'Y': ['* Y Cristal, quarz, oscillators*'],
    'U': ['* U IC *']
}


def write_xls(header, items, handler, sheetname="BOM"):

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(handler)
    worksheet = workbook.add_worksheet()

    tot_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'bg_color': 'lime'})

    hdr_fmt = workbook.add_format({'font_size': 12, 'bold': True, 'bg_color': 'cyan'})

    merge_fmt = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'yellow'})

    for n, i in enumerate(header):
        worksheet.write(0, n, i, hdr_fmt)


    l = []
    row = 1
    for key in ORDER_PATTERN:
        m = filter(lambda x: x[0] == key, items)
        m = map(lambda x: x[1:], m)
        if m:
            row += 1
            worksheet.merge_range('A%s:G%s' % (row, row), ORDER_PATTERN_NAMES[key][0], merge_fmt)
        for rows in m:
            for c, col in enumerate(rows):
                if c == 0:
                    worksheet.write(row, 0, col, tot_fmt)
                else:
                    worksheet.write(row, c, col)
            row += 1

    workbook.close()

def read_xls(handler):
    wb = xlrd.open_workbook(handler)
    data = []

    for s in wb.sheets():
        for row in range(s.nrows):
            values = []
            for col in range(s.ncols):
                curr = s.cell(row,col)
                value = ""
                try:
                    value = str(int(curr.value))
                except (TypeError, ValueError):
                    value = unicode(curr.value)

                values.append(value)
            data.append(values)

        return wb, data

import glob

if __name__ == "__main__":
    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-o", "--out-filename", dest="out_filename", default='merged_bom.xlsx', help="Out file name")
    parser.add_option("-d", "--dirname", dest="dir_name", default='.', help="BOM directory's")
    (options, args) = parser.parse_args()
    print args

    if len(sys.argv) < 2:
        print sys.argv[0], " <xls file name1> <xls file name2> .."
        exit (1)

    file_list = args
    if options.dir_name and not args:
        file_list = glob.glob(os.path.join(options.dir_name, '*.xls'))
        file_list += glob.glob(os.path.join(options.dir_name, '*.xlsx'))

    print "file list........", file_list
    header, data = parse_data(file_list)
    write_xls(header, data, options.out_filename)

