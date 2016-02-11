#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
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

import sys
import os
import re
import xlrd
import xlsxwriter
import datetime

from termcolor import colored

logo_simple = """

#     #                             ######
##   ## ###### #####   ####  ###### #     #  ####  #    #
# # # # #      #    # #    # #      #     # #    # ##  ##
#  #  # #####  #    # #      #####  ######  #    # # ## #
#     # #      #####  #  ### #      #     # #    # #    #
#     # #      #   #  #    # #      #     # #    # #    #
#     # ###### #    #  ####  ###### ######   ####  #    #

"""

logo = """
███╗   ███╗███████╗██████╗  ██████╗ ███████╗██████╗  ██████╗ ███╗   ███╗
████╗ ████║██╔════╝██╔══██╗██╔════╝ ██╔════╝██╔══██╗██╔═══██╗████╗ ████║
██╔████╔██║█████╗  ██████╔╝██║  ███╗█████╗  ██████╔╝██║   ██║██╔████╔██║
██║╚██╔╝██║██╔══╝  ██╔══██╗██║   ██║██╔══╝  ██╔══██╗██║   ██║██║╚██╔╝██║
██║ ╚═╝ ██║███████╗██║  ██║╚██████╔╝███████╗██████╔╝╚██████╔╝██║ ╚═╝ ██║
╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═════╝  ╚═════╝ ╚═╝     ╚═╝

"""

def printout(s, handler, prefix="> ", terminal=True, color='green'):
    s = "%s %s\n" % (prefix, s)
    if terminal:
        s = colored(s, color)

    handler.write(s)
    handler.flush()


def warning(s, handler, prefix=">> ", terminal=True):
    printout(s, handler, prefix=prefix, terminal=terminal, color='yellow')


def error(s, handler, prefix="!! ", terminal=True):
    printout(s, handler, prefix=prefix, terminal=terminal, color='red')


def info(s, handler, prefix="> ", terminal=True):
    printout(s, handler, prefix=prefix, terminal=terminal, color='green')


def order_designator(ref_str):
    ref_str = ref_str.replace(" ", "")
    l = ref_str.split(",")
    try:
        d = sorted(l, key=lambda x: int(re.search('[0-9]+', x).group()))
    except TypeError:
        error("Could not order Designators [%s]" % l)
        sys.exit(1)
    return ", ".join(d)

MULT = {
    'R': 1,
    'k': 1e3,
    'M': 1e6,
    'p': 1e-12,
    'n': 1e-9,
    'u': 1e-6,
}

def order_value(l, handler=sys.stdout, terminal=True):
    if type(l) != tuple and type(l) != list:
        warning("Type data is not a listo or a tuple",
                 handler, terminal=terminal)
        l = [l]

    data = []
    for i in l:
        mult = 1
        v = ""
        accumulator = 0
        # search first multiplier letter
        for n, c in enumerate(i):
            # Skip unit letter
            if c in ["F", "H"]:
                continue
            # Found multiplier convert to number
            if c in MULT:
                accumulator = float(v) * MULT[c]
                mult = MULT[c] / 10.0
                v = ""
                continue
            v += c
        if v:
            try:
                accumulator += float(v) * mult
            except ValueError:
                error("Wrong multiplier [%s]" % c,
                      handler, terminal=terminal)
                sys.exit(1)

        data.append(accumulator)

    return sorted(data)

import math
def eng_string(x):
    '''
    Returns float/int value <x> formatted in a simplified engineering format -
    using an exponent that is a multiple of 3.

    format: printf-style string used to format the value before the exponent.
          1230.0 => 1.23k
      -1230000.0 => -1.23M
    '''
    x = float(x)
    sign = ''
    if x < 0:
        x = -x
        sign = '-'
    exp = int(math.floor( math.log10( x)))
    exp3 = exp - (exp % 3)
    x3 = x / (10 ** exp3)

    if exp3 >= -24 and exp3 <= 24 and exp3 != 0:
        exp3_text = 'yzafpnum kMGTPEZY'[(exp3 - (-24)) / 3]
    elif exp3 == 0:
        exp3_text = ''

    return (sign, str(x3), exp3_text)

def value_toStr(l, handler=sys.stdout, terminal=True):
    data = []
    for i in l:
        value, unit = i
        sign, number, notation = eng_string(value)

        if notation == "":
            number = number.rstrip("0")
            if unit == "ohm":
                number = str(value)
                unit = "R"

        elif notation in ["k","M","G","T","P","E","Z","Y"]:
            number = number.replace(".", notation)
            number = number.rstrip("0")

            if unit == "ohm":
                unit = ""

        elif notation in ["y","z","a","f","p","n","u","m"]:
            number = number.rstrip("0")
            if notation == "m" and unit == "ohm":
                number = str(value)
                unit = "R"
            else:
                number = re.sub(r"\.$", "", number)
                number = "%s%s" % (number, notation)

        number = "%s%s%s" % (sign, number, unit)
        data.append(number)

    return data

# Exchange data layout after file import
FILENAME    = 0
QUANTITY    = 1
DESIGNATOR  = 2
DESCRIPTION = 3
COMMENT     = 4
FOOTPRINT   = 5

VALID_KEYS = [
    u'designator',
    u'comment',
    u'footprint',
    u'description',
]

EXTRA_KEYS = [
    u'date',
    u'project',
    u'hardware_version',
    u'pcb_version',
]

CON = 'J'
CAP = 'C'

VALID_GROUP_KEY = [
    CON,
    'S',
    'F',
    'R',
    CAP,
    'D',
    'DZ',
    'L',
    'Q',
    'TR',
    'Y',
    'U'
]

CATEGORY_NAMES = {
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

NOT_POPULATE_KEY = ["NP", "NM"]


class MergeBom (object):
    def __init__(self, list_bom_files, handler=sys.stdout, terminal=True):
        """
        Data structure

        [
            file_name: [
                designator : [ all ]
                designator : [ all ]
                ...
                designator : [ all ]
            ],
            ...
            file_name: [
                designator : [ all ]
                designator : [ all ]
                ...
                designator : [ all ]
            ],
        ]

        """
        self.files = {}
        self.table_list = []
        self.extra_keys = []
        self.stats = {}
        self.handler = handler
        self.terminal = terminal

        self.stats['file_num'] = 0
        for index_file, file_name in enumerate(list_bom_files):
            self.stats['file_num'] += 1
            warning(file_name, self.handler, terminal=self.terminal)
            wb, data = read_xls(file_name)
            n = os.path.basename(file_name)

            if n in self.files:
                n = "%s-%s" % (index_file, n)
            self.files[n] = index_file

            # Get all header keys
            header = {}
            extra_keys = {}
            for row in data:
                for n, item in enumerate(row):

                    # search header to import corretly all data
                    if item.lower() in VALID_KEYS:
                        header[item.lower()] = n

                    # search extra data, like project name and revision
                    k = ""
                    try:
                        k, v = item.lower().split(":")
                    except ValueError:
                        continue
                    if k in EXTRA_KEYS:
                        extra_keys[k] = v.replace(' ', '')

            self.extra_keys.append(extra_keys)

            try:
                designator  = header['designator']
                comment     = header['comment']
                footprint   = header['footprint']
                description = header['description']
            except KeyError, e:
                error("No key header found! [%s]" % e, self.handler,
                      terminal=self.terminal)
                warning("Valid are:", self.handler, terminal=self.terminal)
                for i in VALID_KEYS:
                    warning("i", self.handler, terminal=self.terminal)

                sys.exit(1)

            table_dict = {}
            for row in data:
                # skip header
                if row[designator].lower() == 'designator':
                    continue

                if filter(lambda x: x, row):
                    # Fix designator field, we want all designator separated
                    # by comma followed by space.
                    # In this way excel could resize row cell correctly
                    if re.findall("\S,[\S]+", row[designator]):
                        row[designator] = row[designator].replace(",", ", ")

                    # Explode designator field to have one component for line
                    d = row[designator].split(',')
                    for reference in d:
                        r = reference.replace(' ', '')
                        if r:
                            table_dict[r] = [
                                os.path.basename(file_name),
                                1,  # one item for a row
                                r,  # designator
                                row[description],
                                row[comment],
                                row[footprint]
                            ]

            self.table_list.append(table_dict)

    def table_data(self):
        return self.table_list

    def extra_data(self):
        return self.extra_keys

    def group(self):
        self.grouped_items = {}
        for table_dict in self.table_list:
            for designator in table_dict.keys():
                # Grop found designator componets by its category
                c = re.search('^[a-zA-Z_]{1,3}', designator)
                group_key = ''
                if c is not None:
                    group_key = c.group().upper()
                    # Buttons and spacer
                    if group_key in ['B', 'BT', 'SCR', 'SPA', 'BAT', 'SW']:
                        group_key = 'S'
                    # Fuses
                    if group_key in ['G']:
                        group_key = 'F'
                    # Tranformer
                    if group_key in ['T']:
                        group_key = 'TR'
                    # Resistors, array, etc.
                    if group_key in ['RN', 'R_G']:
                        group_key = 'R'
                    # Connectors
                    if group_key in ['X', 'P', 'SIM']:
                        group_key = 'J'
                    # Discarted ref
                    if group_key in ['TP']:
                        warning("WARNING!! KEY SKIPPED [%s]" % group_key,
                                self.handler, terminal=self.terminal)
                        continue

                    if group_key not in VALID_GROUP_KEY:
                        error("GROUP key not FOUND!", self.handler,
                              terminal=self.terminal)
                        error("%s, %s, %s" % (c.group(), designator, table_dict[designator]),
                              self.handler, terminal=self.terminal)
                        sys.exit(1)

                    if self.grouped_items.has_key(group_key):
                        self.grouped_items[group_key].append(table_dict[designator])
                    else:
                        self.grouped_items[group_key] = [table_dict[designator]]
                else:
                    error("GROUP key not FOUND!", self.handler, terminal=self.terminal)
                    error(designator, self.handler, terminal=self.terminal)
                    sys.exit(1)

    def table_grouped(self):
        return self.grouped_items

    def count(self):
        """
        grouped items format
        'U': ['bom_due.xls', 1, u'U3000', u'Lan transformer', u'LAN TRANFORMER WE 7490100111A', u'LAN_TR_1.27MM_SMD']
        'U': ['test.xlsx', 1, u'U2015', u'DC/DC switch converter', u'LM22671MR-ADJ', u'SOIC8_PAD']
        'U': ['test.xlsx', 1, u'U2002', u'Dual Low-Power Operational Amplifier', u'LM2902', u'SOIC14']
        'U': ['bom_uno.xls', 1, u'U7', u'Temperature sensor', u'LM75BIM', u'SOIC8']
        """

        self.table = {}

        # Finally Table layout
        self.TABLE_TOTALQTY    = 0
        self.TABLE_DESIGNATOR  = len(self.files) + 1
        self.TABLE_COMMENT     = self.TABLE_DESIGNATOR + 1
        self.TABLE_FOOTPRINT   = self.TABLE_COMMENT + 1
        self.TABLE_DESCRIPTION = self.TABLE_FOOTPRINT + 1

        self.stats['total'] = 0
        for category in VALID_GROUP_KEY:
            if self.grouped_items.has_key(category):
                tmp = {}
                self.stats[category] = 0
                for item in self.grouped_items[category]:
                    if category  == 'J':
                        # Avoid merging for NP componets
                        skip_merge = False
                        for rexp in NOT_POPULATE_KEY:
                            m = re.findall(rexp, item[COMMENT])
                            if m:
                                skip_merge = True
                                error("Not Populate connector, leave unmerged..[%s] [%s] match%s" %
                                       (item[COMMENT], item[DESIGNATOR], m),
                                       self.handler, terminal=self.terminal)
                                item[COMMENT] = item[COMMENT].replace(rexp, ' NP ')

                        if skip_merge:
                            key = item[DESCRIPTION] + item[COMMENT] + item[FOOTPRINT]
                        else:
                            key = item[DESCRIPTION] + item[FOOTPRINT]
                            item[COMMENT] = "Connector"

                        warning("Merged key: %s (%s)" % (key, item[COMMENT]),
                                self.handler, terminal=self.terminal)

                    if category  == 'D' and "LED" in item[FOOTPRINT]:
                            key = item[DESCRIPTION] + item[FOOTPRINT]
                            warning("Merged key: %s (%s)" % (key, item[COMMENT]),
                                    self.handler, terminal=self.terminal)
                    else:
                        key = item[DESCRIPTION] + item[COMMENT] + item[FOOTPRINT]

                    #print key
                    #print "<<", item[DESIGNATOR]
                    self.stats[category] += 1
                    self.stats['total'] += 1

                    # First colum is total Qty
                    curr_file_index = self.files[item[FILENAME]] + self.TABLE_TOTALQTY + 1

                    if key in tmp:
                        #print "UPD", tmp[key], curr_file_index, item[FILENAME]
                        tmp[key][self.TABLE_TOTALQTY] += item[QUANTITY]
                        tmp[key][curr_file_index] += item[QUANTITY]
                        tmp[key][self.TABLE_DESIGNATOR] += ", " + item[DESIGNATOR]
                        tmp[key][self.TABLE_DESIGNATOR] = order_designator(tmp[key][self.TABLE_DESIGNATOR])
                    else:
                        row = [item[QUANTITY]] + \
                              [0] * len(self.files) + \
                              [
                                item[DESIGNATOR],
                                item[COMMENT],
                                item[FOOTPRINT],
                                item[DESCRIPTION]
                              ]

                        row[curr_file_index] = item[QUANTITY]
                        tmp[key] = row
                        #print "NEW", tmp[key], curr_file_index, item[FILENAME]

                self.table[category] = tmp.values()

    def statistics(self):
        return self.stats

    def merge(self):
        self.group()
        self.count()
        for category in VALID_GROUP_KEY:
            if self.table.has_key(category):
                for n, item in enumerate(self.table[category]):
                    self.table[category][n][self.TABLE_DESIGNATOR] = \
                            order_designator(item[self.TABLE_DESIGNATOR])

        return self.table

    def diff(self):
        if len(self.table_list) > 2:
            error("To much file ti compare!", self.handler, terminal=self.terminal)
            sys.exit(1)
        diff = {}
        warning("%s" % self.files.items(), self.handler, terminal=self.terminal)
        for i in self.files.items():
            if i[1] == 0:
                fA = i[0]
                A = self.table_list[0]
            if i[1] == 1:
                fB = i[0]
                B = self.table_list[1]

        warning("A:%s B:%s" % (fA, fB), self.handler, terminal=self.terminal)

        for k in A.keys():
            if B.has_key(k):

                c = re.search('^[a-zA-Z_]{1,3}', A[k][DESIGNATOR])
                category = ''
                if c is not None:
                    category = c.group().upper()

                if category  == 'J':
                    la = [ A[k][DESIGNATOR], A[k][FOOTPRINT] ]
                    lb = [ B[k][DESIGNATOR], B[k][FOOTPRINT] ]

                    warning("Merged key: %s (%s)" % (k, A[k][COMMENT]), self.handler, terminal=self.terminal)

                if category  == 'D' and "LED" in A[k][FOOTPRINT]:
                    la = [ A[k][DESIGNATOR], A[k][FOOTPRINT] ]
                    lb = [ B[k][DESIGNATOR], B[k][FOOTPRINT] ]

                    warning("Merged key: %s (%s)" % (k, A[k][COMMENT]), self.handler, terminal=self.terminal)
                else:
                    la = A[k][1:]
                    lb = B[k][1:]

                if la != lb:
                    diff[k] = (A[k], B[k])

                del B[k]
            else:
                diff[k] = (A[k], [fB] + ['-'] * (len(A[k]) - 1))

            del A[k]

        for k in B.keys():
            diff[k] = ([fA] + ['-'] * (len(B[k]) - 1), B[k])

        return diff

def write_xls(items, file_list, handler, sheetname="BOM", hw_ver="0", pcb_ver="A", project="MyProject",
              diff=False, extra_data=[], statistics=[]):
    STR_ROW = 1
    HDR_ROW = 0
    STR_COL = 0

    A_BOM= "OLD << "
    B_BOM= "NEW >> "

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

    info_fmt_red = workbook.add_format({
        'bold': True,
        'font_color': 'red',
        'font_size':11,
        'valign': 'vcenter',
        'align': 'left',})

    np_fmt = workbook.add_format({
        'bold': True,
        'font_color': 'red',
        'valign': 'vcenter', })

    diffa_fmt = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': 'FFCC33'})
    diffb_fmt = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'fg_color': '#CCFFCC'})
    diff_sep_fmt = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#DDDDDD'})


    hdr_fmt = workbook.add_format({'font_size': 12, 'bold': True, 'bg_color': 'cyan'})
    merge_fmt = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'yellow'})

    # Header info row
    dt = datetime.datetime.now()
    info = []
    for i in file_list:
        info.append("- %s" % i)
    if diff:
        info = [
            'Component Variation',
            '',
            'Date: %s' % dt.strftime("%A, %d %B %Y %X"),
            '',
            '',
            'Project: %s' % project,
            '',
            'Old Revision: %s' % extra_data[0]['revision'],
            'New Revision: %s' % extra_data[1]['revision'],
            '',
            'BOM files:',
        ]
    else:
        info = [
            'Bill of Materials',
            '',
            'Date: %s' % dt.strftime("%A, %d %B %Y %X"),
            '',
            '',
            'Project: %s' % project,
            'Hardware_version: %s' % hw_ver,
            'PCB_version: %s' % pcb_ver,
            '',
            'BOM files:',
        ]

        for i in file_list:
            info.append("- %s" % i)



        # Compute colum len to merge for header
        #stop_col = 'F'
        stop_col = chr(ord('A') + len(file_list+VALID_KEYS))

    row = STR_ROW
    for i in info:
        worksheet.merge_range('A%s:%s%s' % (row, stop_col, row), i, info_fmt)
        row += 1
    row += 1

    # Note and statistics
    worksheet.write('A%s:%s%s' % (row, stop_col, row), "NP=NON MONTARE!", info_fmt_red)
    row += 1

    worksheet.write('A%s:%s%s' % (row, stop_col, row), "Statistics:", info_fmt)
    for i in statistics:
        for c, col in enumerate(i):
            worksheet.write(row, c, col, info_fmt)
        row += 1

    row += 1

    col = 0
    worksheet.write(row, col, "T.Qty", hdr_fmt)
    # fisrt colum is for total quantity.
    col += 1
    if diff:
        sa = "%s [%s]" % (A_BOM, file_list[0])
        sb = "%s [%s]" % (B_BOM, file_list[1])
        worksheet.merge_range('A%s:D%s' % (row, row), sa, diffa_fmt)
        row += 1
        worksheet.merge_range('A%s:D%s' % (row, row), sb, diffb_fmt)
        row += 1

        for i in ['Reference', 'Status', 'Revision', '', 'Reference']:
            worksheet.write(row, col, i.capitalize(), hdr_fmt)
            col += 1
        row += 1
    else:
        for i in file_list:
            worksheet.write(row, col, i, hdr_fmt)
            col += 1

        for i in VALID_KEYS:
            worksheet.write(row, col, i.capitalize(), hdr_fmt)
            col += 1
        row += 1



    row = HDR_ROW + row + 2
    if diff:
        for i in items.keys():
            worksheet.merge_range('A%s:J%s' % (row, row), "%s" % row, diff_sep_fmt)
            A = [i, A_BOM, extra_data[0]['revision'].upper()] + items[i][0][2:]
            B = [i, B_BOM, extra_data[1]['revision'].upper()] + items[i][1][2:]
            error("%s %s %s" % (i, A_BOM, A), self.handler, terminal=self.terminal)
            warning("%s %s %s" % (i, B_BOM, B), self.handler, terminal=self.terminal)
            info("~" * 80, self.handler, terminal=self.terminal, prefix="")

            for n, a in enumerate(A):
                worksheet.write(row,  n, a, diffa_fmt)
                worksheet.write((row + 1), n, B[n], diffb_fmt)

            row += 4
    else:
        l = []

        # Start to write components on xlsx
        for key in VALID_GROUP_KEY:
            if items.has_key(key):
                row += 1
                worksheet.merge_range('A%s:%s%s' % (row, stop_col, row),
                                      CATEGORY_NAMES[key], merge_fmt)
                for i in items[key]:
                    for c, col in enumerate(i):
                        if c == 0:
                            worksheet.write(row, STR_COL, col, tot_fmt)
                        else:
                            # Mark NP to help user
                            fmt = def_fmt
                            if type(col) != int and re.findall("NP", col):
                                    fmt = np_fmt
                            worksheet.write(row, c, col, fmt)
                            if type(col) != int:
                                worksheet.set_column(row, c, len(col) * 6)
                    row += 1

    workbook.close()

def read_xls(handler):
    wb = xlrd.open_workbook(handler)
    data = []

    for s in wb.sheets():
        for row in range(s.nrows):
            values = []
            for col in range(s.ncols):
                try:
                    curr = s.cell(row, col)
                except IndexError:
                    continue

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
    parser.add_option("-p", "--dirname", dest="dir_name", default='.', help="BOM directory's")
    parser.add_option("-d", "--diff", dest="diff", action="store_true", default=False, help="BOM directory's")
    parser.add_option("-r", "--revision", dest="rev", default='0', help="HW Revision")
    parser.add_option("-w", "--pcb-revision", dest="pcb_ver", default='0', help="PCB Revision")
    parser.add_option("-n", "--prj-name", type="string", dest="prj_name", default='MyProject', help="Project names")
    (options, args) = parser.parse_args()
    print args

    if len(sys.argv) < 2:
        print sys.argv[0], " <xls file name1> <xls file name2> .."
        exit (1)

    file_list = args
    if options.dir_name and not args:
        file_list = glob.glob(os.path.join(options.dir_name, '*.xls'))
        file_list += glob.glob(os.path.join(options.dir_name, '*.xlsx'))

    info(logo, sys.stdout, terminal=True, prefix="")
    m = MergeBom(file_list)
    file_list = map(os.path.basename, file_list)

    if options.diff:
        d = m.diff()
        l = m.extra_data()
        write_xls(d, file_list, options.out_filename, diff=True, extra_data=l)
    else:
        d = m.merge()
        stats = m.statistics()
        st = []
        for i in stats.keys():
            if i in CATEGORY_NAMES:
                st.append((stats[i], CATEGORY_NAMES[i]))
        st.append((stats['total'], "Total"))

        write_xls(d, file_list, options.out_filename, hw_ver=options.rev,
                  pcb_ver=options.pcb_ver, project=options.prj_name, statistics=st)


        stats = m.statistics()
        warning("File num: %s" % stats['file_num'], sys.stdout, terminal=True)
        for i in stats.keys():
            if i in CATEGORY_NAMES:
                info(CATEGORY_NAMES[i], sys.stdout, terminal=True, prefix="- ")
                info("%5.5s %5.5s" % (i, stats[i]), sys.stdout, terminal=True, prefix="  ")

        warning("Total: %s" % stats['total'], sys.stdout, terminal=True)
