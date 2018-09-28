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
# Copyright 2018 Daniele Basile <asterix24@gmail.com>
#

import sys
import os
import glob
import getopt
import argparse
import ConfigParser
import re
from lib import cfg,lib, report
from mergebom_class import *
from datetime import datetime

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    # Altium Plugin section
    parser.add_argument('-w', '--workspace-file', dest='workspace_file',
                        help='Altium WorkSpace file, where mergebom grab project infos.', default=None)
    parser.add_argument('-prx', '--bom-prj-file-prefix', dest='bom_prefix',
                        help='Default prefix for BOM in Alrium project', default='bom-')
    parser.add_argument('-ptx', '--bom-prj-file-postfix', dest='bom_postfix',
                        help='Default postfix for BOM in Alrium project', default='')
    parser.add_argument("-pdir", "--bom-prj-dir", dest="bom_search_dir",
                        help="Default name of directory in Altium project where search BOM.", default="Assembly")

    # MergeBom Configuration section
    parser.add_argument("-c", "--merge-cfg", dest="merge_cfg",
                        help="Mergebom configuration file", default=None)
    parser.add_argument("-a", "--csv", dest="csv_file", action="store_true",
                        help="Find and merge csv files, by defaul are excel files.", default=False)
    parser.add_argument("-o", "--out-filename", dest="out_filename",
                        help="Out file name", default='merged_bom')

    parser.add_argument("-l","--log-on-file", dest="log_on_file", action="store_true",
                        help="Log all output in file (by default megebom_report.txt)", default=True)
    parser.add_argument("-p", "--working-dir", dest="working_dir",
                        help="Mergebom working directory", default="./")
    parser.add_argument('-t', '--bom-report-date', dest='report_date_timestamp',
                        help='Default date for merged BOM file in format: %%d/%%m/%%y, by default is today()', default=None)


    # BOM default parameter
    parser.add_argument('-pd', '--prj-date', dest='prj_date',
                        help='Project date time release.', default=None)
    parser.add_argument('-n', '--prj-name',  dest='prj_name',
                        help='Short project name', default=None)
    parser.add_argument('-ln', '--prj-name-long', dest='prj_name_long',
                        help='Long project name', default=None)
    parser.add_argument('-hw', '--prj-hw-ver', dest='prj_hw_ver',
                        help='Project hardware version [0, 1, 2, ..]', default=None)
    parser.add_argument('-pv', '--prj-pcb', dest='prj_pcb',
                        help='PCB hardware version [A, B, C, ..]', default=None)
    parser.add_argument('-lic', '--prj-license',  dest='prj_license',
                        help='prj_license', default=None)
    parser.add_argument('-pn', '--prj-pn', dest='prj_pn',
                        help='Project Part Numeber', default=None)
    parser.add_argument('-s', '--prj-status', dest='prj_status',
                        help='Project status [Prototype, Production, ..]', default=None)

    parser.add_argument("-r", "--replace-original", dest="replace_original", action="store_true",
                        help="delete file", default=False)

    # Diff Mode
    parser.add_argument("-d","--diff", dest="diff", action="store_true",
                        help="Generate diff from two specified BOMs", default=False)

    parser.add_argument('file_to_merge', metavar='N', nargs='*',
                        help='List of file to merge.',
                        default=[])
    options = parser.parse_args()

    if len(sys.argv) < 2:
        parser.print_help()
        sys.exit(0)

    if options.report_date_timestamp is not None:
        options.report_date_timestamp = datetime.strptime(options.report_date_timestamp, '%d/%m/%Y')

    logger = report.Report(log_on_file = options.log_on_file,
                           terminal = True,
                           report_date = options.report_date_timestamp)
    logger.write_logo()

    config = cfg.CfgMergeBom(options.merge_cfg, logger=logger)


    dataset_to_merge = []

    if options.file_to_merge != [] and options.workspace_file is None:
        dataset_to_merge = [
            (options.file_to_merge, {})
        ]

    if options.workspace_file is not None:
        file_BOM = cfg.cfg_altiumWorkspace(options.workspace_file,
                                           options.csv_file,
                                           options.bom_search_dir,
                                           logger,
                                           bom_prefix=options.bom_prefix,
                                           bom_postfix=options.bom_postfix)

        if len(file_BOM) < 1:
            logger.error("No BOM files found in Workspace\n")
            sys.exit(1)

        for item in file_BOM:
            parametri_dict = item[1]
            options.prj_date = parametri_dict.get('prj_date', None)
            options.prj_hw_ver = parametri_dict.get('prj_hw_ver',None)
            options.prj_license = parametri_dict.get('prj_license', None)
            options.prj_name = parametri_dict.get('prj_name', None)
            options.prj_name_long = parametri_dict.get('prj_name_long', None)
            options.prj_pcb = parametri_dict.get('prj_pcb', None)
            options.prj_pn = parametri_dict.get('prj_pn', None)
            options.prj_status = parametri_dict.get('prj_status', None)

            dataset_to_merge.append((item[0], item[1]))
            print dataset_to_merge

    if len(dataset_to_merge[0][0]) == 0:
        logger.error("No file to merge\n")
        sys.exit(1)

    if options.prj_hw_ver is None:
        options.out_filename = "%s_merged" % options.out_filename
    else:
        options.out_filename = "%s-R%s" % (options.out_filename, options.prj_hw_ver)

    #if options.finalf:
    #    appo = merge_file_list[0]
    #    appo = appo.split(os.sep)
    #    options.working_dir = os.path.join(*appo[:-1])

    for item in dataset_to_merge:
        m = MergeBom(item[0], config, is_csv=options.csv_file, logger=logger)
        items = m.merge()
        file_list = map(os.path.basename, item[0])
        out_file = os.path.join(options.working_dir, options.out_filename + '.xlsx')
        extra_data = None
        diff_mode = False
        header_data = cfg.VALID_KEYS

        if options.diff:
            logger.info("Diff Mode..\n")

            items = m.diff()
            extra_data = m.extra_data()
            diff_mode = True
            header_data = m.header_data()

        report.write_xls(items,
            file_list,
            config,
            out_file,
            hw_ver=options.prj_hw_ver,
            pcb=options.prj_pcb,
            name=options.prj_name,
            diff=diff_mode,
            extra_data=extra_data,
            headers=header_data)


