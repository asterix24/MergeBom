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

    parser.add_argument("-v", "--version", dest="mergebom_version", action="store_true",
                        help="Program version", default=False)

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
                        help="Out file name", default=None)

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

    if options.mergebom_version:
        print cfg.MERGEBOM_VER
        sys.exit(0)

    if options.report_date_timestamp is not None:
        options.report_date_timestamp = datetime.strptime(options.report_date_timestamp, '%d/%m/%Y')


    logger = report.Report(log_on_file = options.log_on_file,
                           terminal = True,
                           report_date = options.report_date_timestamp)
    logger.write_logo()


    config = cfg.CfgMergeBom(options.merge_cfg, logger=logger)


    # ===== AltiumWorkspace =============
    if options.workspace_file is not None:
        if options.diff:
            log.error("Invalid switch [--diff], could not run diff mode when merge from altium workspase")
            sys.exit(1)

        bom_dataset = cfg.cfg_altiumWorkspace(options.workspace_file,
                                           options.csv_file,
                                           options.bom_search_dir,
                                           logger,
                                           bom_prefix=options.bom_prefix,
                                           bom_postfix=options.bom_postfix)

        if len(bom_dataset) < 1:
            logger.error("No BOM files found in Workspace\n")
            sys.exit(1)

        logger.info("Found following projects:\n")
        bom_count = 0
        for data in bom_dataset:
            if len(data[1]) == 0:
                continue
            bom_count += 1
            logger.info("%s: %s\n" % (data[0], data[1]))

        bom_dataset_merge = []
        if bom_count == 0:
            logger.error("No Bom to merge found..\n")
            sys.exit(1)
        elif bom_count == 1:
            bom_dataset_merge = bom_dataset
        else:
            logger.info("Wich you want merge?\n")
            logger.info("  1 - All\n", prefix="*")
            logger.info("  2 - interactive\n", prefix="*")
            logger.info("  3 - None\n", prefix="*")
            logger.info(" ", prefix=">>")

            ret = raw_input()
            if ret == "1":
                pass
            elif ret == "2":
                for i in bom_dataset:
                    if len(i[1]) == 0:
                        continue

                    logger.info(i[0] + "\n", prefix="== ")
                    logger.info("Keep or remove [K/r]? ", prefix=">> ")
                    ret = raw_input()
                    if ret in ["r", "R"]:
                        logger.info("Removed!\n", prefix="** ")
                        continue
                    else:
                        logger.info("Added!\n", prefix="** ")
                        bom_dataset_merge.append(i)
            else:
                logger.warning("Bye!\n")
                sys.exit(0)

        logger.warning("==== Start merge ===\n", prefix="")
        for item in bom_dataset_merge:
            if len(item) < 1:
                logger.error("Somethings wrong.. no valid dataset")
                sys.exit(1)

            if len(item[1]) == 0:
                continue

            name = item[0]
            bom = item[1]
            param = item[2]

            if len(bom) > 1:
                logger.warning("There are more than one bom to merge, what wolud I do?\n")
                logger.info("Found following files:\n")
                for i in bom:
                    logger.info(i)
                logger.info("Wich BOM file you want merge?\n")
                logger.info("  1 - All\n", prefix="*")
                logger.info("  2 - interactive\n", prefix="*")
                logger.info("  3 - None\n", prefix="*")
                logger.info(" ", prefix=">>")
                ret = raw_input()
                if ret == "2":
                    for i in bom:
                        print "-" * 80
                        logger.info(i)
                        logger.info("Keep or remove [K/r]?\n", prefix=">> ")
                        ret = raw_input()
                        if ret ["r", "R"]:
                            bom.remove(i)

                if ret == "3":
                    sys.exit(0)


            m = MergeBom(bom, config, is_csv=options.csv_file, logger=logger)

            # Compute outfile name
            hw_ver = param.get(cfg.PRJ_HW_VER, None)
            prj_pcb = param.get(cfg.PRJ_PCB, None)
            prj_name_long = param.get(cfg.PRJ_NAME_LONG, "project")
            prj_name = param.get(cfg.PRJ_NAME, "project")
            prj_pn = param.get(cfg.PRJ_PN, "-")

            out_merge_file = cfg.MERGED_FILE_TEMPLATE_HW % (name, hw_ver)
            if hw_ver is None:
                out_merge_file = cfg.MERGED_FILE_TEMPLATE_NOHW %  name

            wk_path = os.path.dirname(options.workspace_file)
            out = os.path.join(wk_path, options.bom_search_dir, name,out_merge_file)
            report.write_xls(m.merge(),
                map(os.path.basename, bom),
                config,
                out,
                hw_ver=hw_ver,
                pcb=prj_pcb,
                name=prj_name_long + " PN: " + prj_pn,
                diff=False,
                extra_data=None,
                headers=cfg.VALID_KEYS)

            if options.replace_original:
                for i in bom:
                    logger.warning("Remove original..[%s]\n" % i)
                    os.remove(i)

        logger.warning("==== Merge Done! ===\n", prefix="")
        sys.exit(0)

    if len(options.file_to_merge) == 0:
        log.error("You should specify a file to merge or diff.")
        sys.exit(1)

    if options.out_filename is None:
        options.out_filename = "%smerged.xlsx" % options.bom_prefix
        if options.prj_hw_ver is not None:
            if "bom-" == options.bom_prefix:
                options.out_filename = "%smerged-R%s.xlsx" % (options.bom_prefix,
                                                               options.prj_hw_ver)
            else:
                options.out_filename = "%s_merged-R%s.xlsx" % (options.bom_prefix,
                                                               options.prj_hw_ver)

    if options.prj_hw_ver is None:
        logger.warning("Could be nice to specify a hardware version for better tracking.")

    if options.prj_name is None:
        logger.warning("Could be nice to specify a project name for better tracking.")



    m = MergeBom(options.file_to_merge, config, is_csv=options.csv_file, logger=logger)

    items = m.merge()
    file_list = map(os.path.basename, options.file_to_merge)
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
        os.path.join(options.working_dir, options.out_filename),
        hw_ver=options.prj_hw_ver,
        pcb=options.prj_pcb,
        name=options.prj_name,
        diff=diff_mode,
        extra_data=extra_data,
        headers=header_data)


