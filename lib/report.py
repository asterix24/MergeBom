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
# Copyright 2017 Daniele Basile <asterix24@gmail.com>
#

import sys
import os
import re
import xlrd
import xlsxwriter
import datetime

from lib import *

def write_xls(items, file_list, cfg, handler, sheetname="BOM", hw_ver="0", pcb_ver="A", project="MyProject",
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
    worksheet.write('A%s:%s%s' % (row, stop_col, row), "NP=NOT POPULATE (NON MONTARE)!", info_fmt_red)
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
        categories = cfg.getCategories()
        for key in categories:
            if items.has_key(key):
                row += 1
                title = "%s * %s *" % (key, cfg.get(key, 'desc'))
                worksheet.merge_range('A%s:%s%s' % (row, stop_col, row),
                                      title, merge_fmt)
                for i in items[key]:
                    for c, col in enumerate(i):
                        if c == 0:
                            worksheet.write(row, STR_COL, col, tot_fmt)
                        else:
                            # Mark NP to help user
                            fmt = def_fmt
                            if type(col) != int and re.findall(NP_REGEXP, col):
                                    fmt = np_fmt
                            worksheet.write(row, c, col, fmt)
                            if type(col) != int:
                                worksheet.set_column(row, c, 50)
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


