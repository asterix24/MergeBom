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

"""
MergeBOM Report module.
Utils to generate and read excel BOM files.
"""

import os
import sys
import re
import datetime
import xlrd
import xlsxwriter
import lib
import cfg
import csv
from termcolor import colored



class Report(object):
    """
    Merge Bom report generator.
    """

    def __init__(self, logfile="./mergebom_report.txt", log_on_file=False,
                 terminal=True, report_date=None):
        self.terminal = terminal
        self.report = None
        self.log_on_file = log_on_file
        self.report_date=report_date
        if report_date is None:
            self.report_date = datetime.datetime.now()
        if self.log_on_file:
            self.report = open(logfile, 'w+')

    def __del__(self):
        if self.log_on_file:
            self.report.flush()
            self.report.close()

    def __printout(self, s, prefix="", color=""):
        s = "%s%s" % (prefix, s)
        if self.terminal:
            out = s
            if color:
                out = colored(out, color)
            sys.stdout.write(out)
            sys.stdout.flush()

        if self.log_on_file:
            self.report.write(s)
            self.report.flush()

    def write_logo(self):
        logo = cfg.LOGO_SIMPLE
        if self.terminal and not self.log_on_file:
            logo = cfg.LOGO
        self.__printout(logo, color='green')

    def write_header(self, conf_key, file_list):
        """
        Write BOM Header info in report file.
        """

        self.__printout("Report file.\n")
        self.__printout("MergeBom Version: %s\n" % cfg.MERGEBOM_VER)

        
        self.__printout("Date: %s\n" % self.report_date.strftime('%d/%m/%Y'))
        self.__printout("." * 80)
        self.__printout("\n")
        self.__printout("\n")
        self.__printout(":" * 80)
        self.__printout("Date: %s\n" % conf_key.get('date', '-'))
        self.__printout("Project Name: %s\n" % conf_key.get('name','-'))
        self.__printout("Hardware Revision: %s\n" % conf_key.get('hw_ver','-'))
        self.__printout("PCB Revision: %s\n" % conf_key.get('pcb_ver','-'))
        self.__printout("\n")

        self.__printout("Bom Files:\n")
        for i in file_list:
            self.__printout(" - %s\n" % i)

        self.__printout("\n== Check Merged items: ==\n")
        self.__printout("-" * 80)
        self.__printout("\n")

    def write_stats(self, stats):
        """
        Write BOM component Statistics
        """
        self.__printout("Total: %s\n" % stats['total'])

    def warning(self, s, prefix=">> "):
        self.__printout(s, prefix=prefix, color='yellow')


    def error(self, s, prefix="!! "):
        self.__printout(s, prefix=prefix, color='red')


    def info(self, s, prefix="> "):
        self.__printout(s, prefix=prefix, color='green')



def write_xls(
        items,
        file_list,
        config,
        handler,
        hw_ver="0",
        pcb="A",
        name="MyProject",
        diff=False,
        extra_data=None,
        statistics=None,
        headers=cfg.VALID_KEYS):
    """
    Write merged BOM in excel file.
    Statistics data should be in follow format:
    {'total': 38, 'R': 14, 'file_num': 2, 'C': 12}
    Where:
    - total: sum all of components numeber
    - file_num: number of merged BOM files
    - C,R, J, ecc.: Componentes category
    """

    STR_ROW = 1
    HDR_ROW = 0
    STR_COL = 0

    A_BOM = "FILE A << "
    B_BOM = "FILE B >> "

    if diff:
        if extra_data is not None:
            A_hw_diff = "file-a"
            B_hw_diff = "file-b"
            A_pcb_diff = "file-a"
            B_pcb_diff = "file-b"

            if len(extra_data) > 1:
                A_hw_diff = extra_data[0].get("hardware_version", "-")
                B_hw_diff = extra_data[1].get("hardware_version", "-")
                A_pcb_diff = extra_data[0].get("pcb_version", "-")
                B_pcb_diff = extra_data[1].get("pcb_version", "-")

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(handler)
    worksheet = workbook.add_worksheet()

    def_fmt = workbook.add_format({'valign': 'vcenter'})
    def_fmt.set_text_wrap()

    stat_hdr_fmt = workbook.add_format({
        'bold': True,
        'align': 'left',
        'valign': 'vcenter',
        'bg_color': '#66B2FF'})

    stat_hdr1_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#00CCCC'})

    stat_num_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': 'silver'})

    stat_desc_fmt = workbook.add_format({
        'italic': True,
        'align': 'left',
        'valign': 'vcenter' })

    stat_ctot_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#4C9900'})

    stat_ltot_fmt = workbook.add_format({
        'bold': True,
        'align': 'left',
        'valign': 'vcenter',
        'bg_color': '#4C9900'})

    tot_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': 'lime'})

    info_fmt = workbook.add_format({
        'bold': True,
        'font_size': 11,
        'valign': 'vcenter',
        'align': 'left', })

    info_fmt_red = workbook.add_format({
        'bold': True,
        'font_color': 'red',
        'font_size': 11,
        'valign': 'vcenter',
        'align': 'left', })

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

    hdr_fmt = workbook.add_format(
        {'font_size': 12, 'bold': True, 'bg_color': 'cyan'})
    merge_fmt = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'yellow'})

    # Header info row
    report_date = datetime.datetime.now()
    info = []
    for i in file_list:
        info.append("- %s" % i)
    if diff:
        info = [
            'Component Variation',
            '',
            'Date: %s' % report_date.strftime('%d/%m/%Y'),
            '',
            '',
            'Project: %s' % name,
            '',
            'File A:',
            '   Hw_rev:  %s' % A_hw_diff,
            '   PCB_rev: %s' % A_pcb_diff,
            'File B:',
            '   Hw_rev:  %s' % B_hw_diff,
            '   PCB_rev: %s' % B_pcb_diff,
            '',
            'BOM files:',
        ]
    else:
        info = [
            'Bill of Materials',
            '',
            'Date: %s' % report_date.strftime('%d/%m/%Y'),
            '',
            '',
            'Project: %s' % name,
            'Hardware_version: %s' % hw_ver,
            'PCB_version: %s' % pcb,
            '',
            'BOM files:',
        ]

        for i in file_list:
            info.append("- %s" % i)

    # Compute colum len to merge for header
    #stop_col = 'F'
    stop_col = chr(ord('A') + len(file_list + headers))

    row = STR_ROW
    for i in info:
        worksheet.merge_range('A%s:%s%s' % (row, stop_col, row), i, info_fmt)
        row += 1
    row += 1

    # Statistics
    worksheet.merge_range('A%s:%s%s' % (row, 'C', row), "Statistics", stat_hdr1_fmt)
    categories = config.categories()
    if statistics is not None:
        # Write description:
        worksheet.write(row, 0, statistics.get("file_num", "-"), stat_num_fmt)
        row += 1
        worksheet.merge_range('B%s:C%s' % (row, row), "Merged Files Number", stat_desc_fmt)
        row += 1
        worksheet.merge_range('A%s:C%s' % (row, row), "Components:", stat_hdr_fmt)
        for cat in categories:
            if cat in statistics:
                worksheet.write(row, 0, statistics[cat], stat_num_fmt)
                row += 1
                worksheet.merge_range('B%s:C%s' % (row, row),
                                config.get(cat, "desc"), stat_desc_fmt)

        worksheet.write(row, 0, statistics.get("total", "-"), stat_ctot_fmt)
        row += 1
        worksheet.merge_range('B%s:C%s' % (row, row), "Total BOM Componets", stat_ltot_fmt)
        row += 1

    row += 2
    # Note
    worksheet.write('A%s:%s%s' % (row, stop_col, row),
                    "NP=NOT POPULATE (NON MONTARE)!", info_fmt_red)
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

        for i in headers:
            worksheet.write(row, col, i.capitalize(), hdr_fmt)
            col += 1
        row += 1

    row = HDR_ROW + row + 2
    if diff:
        for i in items.keys():
            worksheet.merge_range(
                'A%s:J%s' %
                (row, row), "%s" %
                row, diff_sep_fmt)
            A = [i, A_BOM, A_hw_diff.upper()] + items[i][0][2:]
            B = [i, B_BOM, B_hw_diff.upper()] + items[i][1][2:]

            for n, a in enumerate(A):
                worksheet.write(row, n, a, diffa_fmt)
                worksheet.write((row + 1), n, B[n], diffb_fmt)

            row += 4
    else:
        # Start to write components on xlsx
        categories = config.categories()
        for key in categories:
            if key in items:
                row += 1
                title = "%s * %s *" % (key, config.get(key, 'desc'))
                worksheet.merge_range('A%s:%s%s' % (row, stop_col, row),
                                      title, merge_fmt)
                for i in items[key]:
                    for c, col in enumerate(i):
                        if c == 0:
                            worksheet.write(row, STR_COL, col, tot_fmt)
                        else:
                            # Mark NP to help user
                            fmt = def_fmt
                            if not isinstance(
                                    col, int) and re.findall(
                                        cfg.NP_REGEXP, col):
                                fmt = np_fmt
                            worksheet.write(row, c, col, fmt)
                            if not isinstance(col, int):
                                worksheet.set_column(row, c, 50)
                    row += 1

    workbook.close()

import csv, codecs

class UTF8Recoder:
    """
    Iterator that reads an encoded stream and reencodes the input to UTF-8
    """
    def __init__(self, f, encoding):
        self.reader = codecs.getreader(encoding)(f)

    def __iter__(self):
        return self

    def next(self):
        return self.reader.next().encode("utf-8")


class DataReader(object):
    """
    Read data from BOM.
    """
    def __init__(self, file_name, is_csv=False):
        self.is_csv = is_csv
        self.file_name = file_name

    def __xls_reader(self):
        wb = xlrd.open_workbook(self.file_name)
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

            return data

    def __csv_reader(self):
        data = []
        f = UTF8Recoder(open(self.file_name, 'rb'), "cp1252")
        rd = csv.reader(f, delimiter=',', quotechar='\"')
        for item in rd:
            if item:
                data.append([unicode(s, "cp1252") for s in item])
        return data

    def read(self):
        if self.is_csv:
            return self.__csv_reader()
        else:
            return self.__xls_reader()

if __name__ == "__main__":

    config = cfg.CfgMergeBom()
    rep = Report(log_on_file=True)
    rep.write_logo()
    rep.write_header({},[])
    rep.error("Errore..\n")
    rep.info("Info..\n")
    rep.warning("Warning..\n")

