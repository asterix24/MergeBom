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
import xlwt
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
TOT_COUNT = 0

def parse_data(bom_file_list):
    FILES=len(bom_file_list)
    table_dict = {}

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

                raw_row = [''] * (FILES + TOT_COUNT + 5)

                if table_dict.has_key(key):
                    raw_row[TOT_COUNT] = int(j[QUANTITY]) + table_dict[key][TOT_COUNT]
                else:
                    raw_row[TOT_COUNT] = int(j[QUANTITY])

                raw_row[index_file + 2] = int(j[QUANTITY])
                raw_row[index_file + 2 + DESIGNATOR] = j[QUANTITY]
                raw_row[index_file + 2 + COMMENT] = j[COMMENT]
                raw_row[index_file + 2 + FOOTPRINT] = j[FOOTPRINT]
                raw_row[index_file + 2 + DESCRIPTION] = j[DESCRIPTION]

                table_dict[key] = raw_row

    return table_dict

def foo():
    d = {}
    l = sorted(table_dict.values(), key=lambda ref: ref[DESCRIPTION][:2])
    for g in l:
        if re.findall("\S,[\S]+", g[DESCRIPTION]):
            g[DESCRIPTION] = g[DESCRIPTION].replace(",", ", ")
            print "........", g[DESCRIPTION].replace(",", ", ")
        c = re.search('^[a-zA-Z_]{1,3}', g[DESCRIPTION])
        key = c.group().upper()

        # Buttons and spacer
        if key in ['B', 'BT', 'SCR', 'SPA', 'BAT','SW']:
            key = 'S'
        # Fuses
        if key in ['G']:
            key = 'F'
        # Tranformer
        if key in ['T' ]:
            key = 'TR'
        # Resistors, array, etc.
        if key in ['RN', 'R_G']:
            key = 'R'
        # Connectors
        if key in ['X', 'P', 'SIM']:
            key = 'J'
        # Discarted ref
        if key in ['TP']:
            print "WARNING WE SKIP THIS KEY"
            print 'key [%s]' % key
            print '.........'
            continue

        if d.has_key(key):
            d[key].append(g)
        else:
            d[key] = [g]

        print 'Group Key:',c.group(), g[DESCRIPTION]

    SEPARATOR_NUM = len(l[0]) - 1
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
    #Add separator from each group of components.
    for i in ORDER_PATTERN_NAMES.keys():
        ORDER_PATTERN_NAMES[i] = (([''] * DESCRIPTION) + ORDER_PATTERN_NAMES[i] + ([''] * (SEPARATOR_NUM - DESCRIPTION)))
        print ORDER_PATTERN_NAMES[i]


    for k in d.keys():
        if k in ['D', 'J', 'S', 'U']:
            d[k] = sorted(d[k], key=lambda ref: ref[DESCRIPTION_PLUS])
        else:
            d[k] = sorted(d[k], key=lambda ref: ref[COMMENT_PLUS])

    #Check missing group code.
    for j in d.keys():
        if j not in ORDER_PATTERN:
            print 'Missing order pattern key: \"%s\"' % j
            print 'In BOM:', d.keys()
            print 'In mergebom:', ORDER_PATTERN
            sys.exit(0)

    with open(out_filename, 'wb') as csvfile:
        data = csv.writer(csvfile, delimiter=out_delimiter)
        data.writerow(header)
        for p in ORDER_PATTERN:
            if d.has_key(p):
                data.writerow(ORDER_PATTERN_NAMES[p])
                for i in d[p]:
                    data.writerow(i)

    print
    print
    print fillRowCenter("=" * 80, " Final Report ")
    print
    print

    total = 0
    recap = {}
    for p in ORDER_PATTERN:
        if d.has_key(p):
            s = ORDER_PATTERN_NAMES[p][2]
            s = s.replace('*','')
            print
            print fillRowCenter("*" * 80, s)
            recap[s] = 0
            for i in d[p]:
                if i[QUANTITY] != '':
                    total += i[QUANTITY]
                    recap[s] += i[QUANTITY]
                    print fillTableRow(" " * 80, "n.%d" % i[QUANTITY], i[COMMENT_PLUS], i[FOOTPRINT_PLUS])


    print
    print
    print "~" * 10, "Total", "~" *10
    for r in recap.keys():
        print "%4d: %s" % (recap[r], r)

    print "-" * 24
    print "%4d Total components" % total
    print "~" * 24
    print
    print


HEADER_ROW = 0
CELL_WIDTH = 4000

import logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.ERROR)

def xldate_as_datetime(book, xldate):
    xldate_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(xldate, book.datemode))
    logger.info('datetime: %s' % xldate_as_datetime)
    return xldate_as_datetime

def write_xls(items, handler, sheetname="Elenco"):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet(sheetname)

    for colum, item in enumerate(items):
        sheet.write(HEADER_ROW, colum, str(item))
        sheet.col(colum).width = CELL_WIDTH

    book.save(handler)
   
def read_xls(handler):
    wb = xlrd.open_workbook(handler)
    data = []

    for s in wb.sheets():
        logger.debug('Sheet:',s.name, s.ncols, s.nrows)
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



if __name__ == "__main__":
    from optparse import OptionParser

    parser = OptionParser()
    (options, args) = parser.parse_args()

    if len(sys.argv) < 2:
        print sys.argv[0], " <xls file name1> <xls file name2> .."
        exit (1)

    out_filename = 'merged_bom.xls'
    print args
    print parse_data(args)

