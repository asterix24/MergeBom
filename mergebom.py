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
    file_list = []
    parser = argparse.ArgumentParser()
    parser.add_argument('--path-workspace-file', '-w', dest='workspace_file',
                        help='Dove si trova il file WorkSpace', default=None)
    parser.add_argument('--bom-prj-file-prefix', '-prefix', dest='bom_prefix',
                        help='Default prefix for BOM in project', default='bom-')
    parser.add_argument('--bom-prj-file-postfix', '-postfix', dest='bom_postfix',
                        help='Default postfix for BOM in project', default='')
    parser.add_argument('--nome-directory-final-file', '-finalf', dest='finalf', action="store_true",
                        help='Se il file deve avere lo stesso nome e la stessa directory del file vecchio', default=False)
    parser.add_argument("-a", "--csv", dest="csv_file", action="store_true",
                        default=False, help="Find and merge csv files, by defaul are excel files.")
    parser.add_argument("-c", "--merge-cfg", dest="merge_cfg",
                        default=None, help="MergeBOM configuration file.")
    parser.add_argument("-o", "--out-filename", dest="out_filename",
                        default='merged_bom', help="Out file name")
    parser.add_argument("-p", "--working-dir", dest="working_dir",
                        default="./", help="BOM to merge working path.")
    parser.add_argument("-s", "--bom-file-search-dir", dest="bom_search_dir",
                        default="Assembly", help="Project BOM search directory.")
    parser.add_argument('--report_time', '-t', dest='report_time',
                        help='datetime nel formato : %%d/%%m/%%y', default=None)
    parser.add_argument("-r", "--bom-revision", dest="bom_rev",
                        default=None, help="Hardware BOM revision")
    parser.add_argument("-pc", "--bom-pcb-revision", dest="bom_pcb_ver",
                        default=None, help="PCB Revision")
    parser.add_argument("-n", "--bom-prj-name", dest="bom_prj_name",
                        default=None, help="Project names.")
    parser.add_argument("-d", "--delete-file", dest="delete",action="store_true",
                        default=False, help="delete file")
    parser.add_argument("-l","--log-on-file",dest="log_on_file",
                        default=True,action="store_true",help="List all project name from version file.")
    parser.add_argument( "-diff","--diff",dest="diff",action="store_true",
                        default=False, help="Generate diff from two specified BOMs")

    parser.add_argument('--prj_date', '-date', dest='prj_date',
                        help='prj_date', default=None)
    parser.add_argument('--prj_hw_ver', '-hw_ver', dest='prj_hw_ver',
                        help='prj_hw_ver', default=None)
    parser.add_argument('--prj_license', '-license', dest='prj_license',
                        help='prj_license', default=None)
    parser.add_argument('--prj_name', '-name', dest='prj_name',
                        help='prj_name', default=None)
    parser.add_argument('--prj_name_long', '-name_long', dest='prj_name_long',
                        help='prj_name_long', default=None)
    parser.add_argument('--prj_pcb', '-pcb', dest='prj_pcb',
                        help='prj_pcb', default=None)
    parser.add_argument('--prj_pn', '-pn', dest='prj_pn',
                        help='prj_pn', default=None)
    parser.add_argument('--prj_status', '-status', dest='prj_status',
                        help='prj_status', default=None)

    parser.add_argument('file_to_merge', metavar='N', nargs='*',
                        help='List of file to merge.',
                        default=[])
    options = parser.parse_args()

    if len(sys.argv) < 2:
        parser.print_help()
        sys.exit(0)

    if options.report_time is not None:
        options.report_time = datetime.strptime(options.report_time, '%d/%m/%Y')

    logger = report.Report(log_on_file = options.log_on_file,
                           terminal = True,
                           report_date = options.report_time)
    logger.write_logo()

    config = cfg.CfgMergeBom(options.merge_cfg, logger=logger)

    if options.finalf:
        options.out_filename = options.name_file

    merge_file_list = options.file_to_merge
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

            for j in item[0]:
                merge_file_list.append(j)

    if len(merge_file_list) == 0:
        logger.error("No file to merge\n")
        sys.exit(1)

    if not options.delete:
        if options.prj_hw_ver is None:
            options.out_filename = options.out_filename + '_merge'
        else:
            options.out_filename = options.out_filename + options.prj_hw_ver

    if options.finalf:
        appo = merge_file_list[0]
        appo = appo.split(os.sep)
        options.working_dir = os.path.join(*appo[:-1])

    m = MergeBom(merge_file_list, config, logger=logger)
    items = m.merge()
    file_list = map(os.path.basename, merge_file_list)
    out_file = os.path.join(options.working_dir, options.out_filename + '.xlsx')
    extra_data = None
    diff_mode = False
    header_data = cfg.VALID_KEYS

    if options.diff:
        logger.info("Diff Mode..\n")
        if len(merge_file_list) != 2:
            logger.error("No enougth file to merge\n")
            sys.exit(1)

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

    if options.delete:
        logger.info("Remove Merged files\n")
        for i,v in enumerate(merge_file_list):
            os.remove(merge_file_list[i])

