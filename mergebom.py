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
import glob
import argparse
import ConfigParser
import re
from lib import cfg
from mergebom_class import *
   
if __name__ == "__main__":

    file_list = []
    parser = argparse.ArgumentParser()
    parser.add_argument('--workspace_file', '-w', dest='ws', 
                        help='Dove si trova il file WorkSpace', default='./test/utils.DsnWrk')
    parser.add_argument("-a", "--csv", dest="csv_file", action="store_true",
                      default=False, help="Find and merge csv files, by defaul are excel files.")
    parser.add_argument("-c", "--merge-cfg", dest="merge_cfg",
                      default=None, help="MergeBOM configuration file.")
    parser.add_argument("-o", "--out-filename", dest="out_filename",
                      default='merged_bom.xlsx', help="Out file name")
    parser.add_argument("-p", "--working-dir", dest="working_dir",
                      default='./', help="BOM to merge working path.")                  
    parser.add_argument(
        "-l",
        "--log-on-file",
        dest="log_on_file",
        default=True,
        action="store_true",
        help="List all project name from version file.")
    options = parser.parse_args()
    file_BOM = cfg.cfg_altiumWorkspace(options.ws, options.csv_file) 
    if not file_BOM:
        print("i file non sono stati trovati")
    else:
        config = cfg.CfgMergeBom(options.merge_cfg)
        logger = report.Report(log_on_file=options.log_on_file, terminal=True)
        logger.write_logo()

        appo = []
        file_list = []
        for i in range(0,len(file_BOM)):
            appo = file_BOM[i][0]
            for j in range(0,len(appo)):
                file_list.append(appo[j])

        m = MergeBom(file_list, config, logger=logger)
        d = m.merge()
        file_list = map(os.path.basename, file_list)
        ft = os.path.join(options.working_dir, options.out_filename)
        report.write_xls(
            d,
            file_list,
            config,
            ft)

