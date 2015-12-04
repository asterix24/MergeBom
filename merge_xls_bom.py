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
# Copyright 2015 Daniele Basile <asterix24@gmail.com>
#
#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import sys, os
import re

import xlrd
import xlsxwriter
import datetime

from termcolor import colored

def warning(s):
    print colored(">> " + s,'yellow')
def error(s):
    print colored("!! " + s,'red')

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

    col1 = col1[:min(len(col1), 10)]
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

FILENAME    = 0
QUANTITY    = 1
DESIGNATOR  = 2
DESCRIPTION = 3
COMMENT     = 4
FOOTPRINT   = 5
FILES = {}

valid_keys = [u'quantity', u'designator', u'comment', u'footprint', u'description', u'libref']

def import_data(bom_file_list):
    table_dict = {}
    header = {}
    global FILES

    for index_file, file_name in enumerate(bom_file_list):
        wb, data = read_xls(file_name)
        FILES[os.path.basename(file_name)] = index_file

        # Get all header keys
        for row in data:
            for n, item in enumerate(row):
                if item.lower() in valid_keys:
                    header[item.lower()] = n

        quantity    = header['quantity']
        designator  = header['designator']
        comment     = header['comment']
        footprint   = header['footprint']
        description = header['description']


        for row in data:
            # skip header
            if row[designator].lower() == 'designator':
                continue

            if filter(lambda x: x, row):
                # Fix designator field, we want all designator separated by
                # comma followed by space. In this way excel could resize row
                # cell correctly
                if re.findall("\S,[\S]+", row[designator]):
                    row[designator] = row[designator].replace(",", ", ")

                # Explode designator field to have one component for line
                d = row[designator].split(',')
                for reference in d:
                    r = reference.replace(' ', '')
                    if r:
                        table_dict[r] = [os.path.basename(file_name), 1, r, row[description],
                            row[comment], row[footprint]]

    return header, table_dict

CAP = 'C'

valid_group_key = ['J', 'S', 'F','R', CAP,'D','DZ','L', 'Q','TR','Y', 'U']

def group_items(table_dict):
    grouped_items = {}
    for designator in table_dict.keys():
        # Grop found designator componets by its category
        c = re.search('^[a-zA-Z_]{1,3}', designator)
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
                warning("WARNING!! KEY SKIPPED [%s]" % group_key)
                continue

            if group_key not in valid_group_key:
                error("GROUP key not FOUND!")
                error("%s, %s, %s" % (c.group(), designator, table_dict[designator]))
                sys.exit(1)

            if grouped_items.has_key(group_key):
                grouped_items[group_key].append(table_dict[designator])
            else:
                grouped_items[group_key] = [table_dict[designator]]
        else:
            error("GROUP key not FOUND!")
            error(designator)
            sys.exit(1)

    return grouped_items

def diff_table(grouped_items):
    diff = {}
    for category in valid_group_key:
        if grouped_items.has_key(category):
            table_a = {}
            table_b = {}
            a = None
            for item in grouped_items[category]:
                print item[FILENAME], item
                key = item[DESCRIPTION] + item[COMMENT] + item[FOOTPRINT]

    return table_a, table_b

ORDER_PATTERN = [
    'J', 'S', 'F','R',
    CAP,
    'D','DZ','L', 'Q','TR','Y', 'U'
]

ORDER_PATTERN_NAMES = {
    'J':  'J  Connectors',
    'S':  'S  Mechanical parts and buttons',
    'F':  'F  Fuses',
    'R':  'R  Resistors',
     CAP :  'C  Capacitors',
    'D':  'D  Diodes',
    'DZ': 'DZ Zener, Schottky, Transil',
    'L':  'L  Inductors, chokes',
    'Q':  'Q  Transistors',
    'TR': 'TR Transformers',
    'Y':  'Y  Cristal, quarz, oscillator',
    'U':  'U  IC',
}


def grouped_count(grouped_items):
    """
    grouped items format
    'U': ['bom_due.xls', 1, u'U3000', u'Lan transformer', u'LAN TRANFORMER WE 7490100111A', u'LAN_TR_1.27MM_SMD']
    'U': ['test.xlsx', 1, u'U2015', u'DC/DC switch converter', u'LM22671MR-ADJ', u'SOIC8_PAD']
    'U': ['test.xlsx', 1, u'U2002', u'Dual Low-Power Operational Amplifier', u'LM2902', u'SOIC14']
    'U': ['bom_uno.xls', 1, u'U7', u'Temperature sensor', u'LM75BIM', u'SOIC8']
    """

    table = {}
    # Finally Table layout
    global TABLE_TOTALQTY
    global TABLE_DESIGNATOR
    global TABLE_COMMENT
    global TABLE_FOOTPRINT
    global TABLE_DESCRIPTION

    TABLE_TOTALQTY    = 0
    TABLE_DESIGNATOR  = len(FILES) + 1
    TABLE_COMMENT     = TABLE_DESIGNATOR + 1
    TABLE_FOOTPRINT   = TABLE_COMMENT + 1
    TABLE_DESCRIPTION = TABLE_FOOTPRINT + 1

    for category in valid_group_key:
        if grouped_items.has_key(category):
            tmp = {}
            count = 0
            for item in grouped_items[category]:
                if category  == 'J':
                    key = item[DESCRIPTION] + item[FOOTPRINT]
                    warning("Merged key: %s (%s)" % (key, item[COMMENT]))
                    item[COMMENT] = "Connector"
                if category  == 'D' and "LED" in item[FOOTPRINT]:
                        key = item[DESCRIPTION] + item[FOOTPRINT]
                        warning("Merged key: %s (%s)" % (key, item[COMMENT]))
                else:
                    key = item[DESCRIPTION] + item[COMMENT] + item[FOOTPRINT]

                #print key
                #print "<<", item[DESIGNATOR]
                count += 1

                # First colum is total Qty
                curr_file_index = FILES[item[FILENAME]] + TABLE_TOTALQTY + 1

                if tmp.has_key(key):
                    #print "UPD", tmp[key], curr_file_index, item[FILENAME]
                    tmp[key][TABLE_TOTALQTY] += item[QUANTITY]
                    tmp[key][curr_file_index] += item[QUANTITY]
                    tmp[key][TABLE_DESIGNATOR] += ", " + item[DESIGNATOR]
                else:
                    row = [item[QUANTITY]] + \
                          [0] * len(FILES) + \
                          [
                            item[DESIGNATOR],
                            item[COMMENT],
                            item[FOOTPRINT],
                            item[DESCRIPTION]
                          ]

                    row[curr_file_index] = item[QUANTITY]
                    tmp[key] = row
                    #print "NEW", tmp[key], curr_file_index, item[FILENAME]
            print count

            table[category] = tmp.values()

    return table


def write_xls(header, items, file_list, handler, sheetname="BOM"):
    STR_ROW = 1
    HDR_ROW = 0
    STR_COL = 0

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(handler)
    worksheet = workbook.add_worksheet()

    def_fmt = workbook.add_format({'valign': 'vcenter'})
    def_fmt.set_text_wrap()

    tot_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': 'lime'})

    info_fmt = workbook.add_format({
        'bold': True,
        'font_size':11,
        'valign': 'vcenter',
        'align': 'left',})

    hdr_fmt = workbook.add_format({'font_size': 12, 'bold': True, 'bg_color': 'cyan'})
    merge_fmt = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'yellow'})

    # Time info
    dt = datetime.datetime.now()
    info = [
        'Bill of Materials',
        '',
        'Date: %s' % dt.strftime("%A, %d %B %Y %X"),
        '',
        '',
        'Project:',
        'Revision:',
        '',
        'BOM files:',
    ]
    for i in file_list:
        info.append("- %s" % i)

    row = STR_ROW
    for i in info:
        worksheet.merge_range('A%s:O%s' % (row, row), i, info_fmt)
        row += 1

    # Header info row
    col = 0
    worksheet.write(row, col, "T.Qty", hdr_fmt)
    col += 1
    for i in file_list:
        worksheet.write(row, col, i, hdr_fmt)
        col += 1

    for i in [u'Designator', u'Comment', u'Footprint', u'Description']:
        worksheet.write(row, col, i, hdr_fmt)
        col += 1
    row += 1

    l = []
    # Start to write components on xlsx
    row = HDR_ROW + row + 2
    for key in ORDER_PATTERN:
        if items.has_key(key):
            row += 1
            worksheet.merge_range('A%s:O%s' % (row, row), ORDER_PATTERN_NAMES[key], merge_fmt)
            for i in items[key]:
                for c, col in enumerate(i):
                    if c == 0:
                        worksheet.write(row, STR_COL, col, tot_fmt)
                    else:
                        worksheet.write(row, c, col, def_fmt)
                row += 1

    workbook.close()

    return {}

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

    header, data = import_data(file_list)
    data = group_items(data)
    data = grouped_count(data)
    file_list = map(os.path.basename, file_list)
    stats = write_xls(header, data, file_list, options.out_filename)

    #print
    #print
    #print ":-" * 40
    #for i in ORDER_PATTERN:
    #    if stats.has_key(i):
    #        print "%5s %s" % (stats[i], ORDER_PATTERN_NAMES[i])

    #print "=" * 80
    #print "Total: %s" % (stats['total'])
    #print ":-" * 40

