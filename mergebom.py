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
import getopt
import argparse
import ConfigParser
import re
from lib import cfg
from mergebom_class import *
from datetime import datetime
   
if __name__ == "__main__":

    file_list = []
    parser = argparse.ArgumentParser()
    parser.add_argument('--workspace_file', '-w', dest='ws', 
                        help='Dove si trova il file WorkSpace', default=None)
    parser.add_argument("-a", "--csv", dest="csv_file", action="store_true",
                      default=False, help="Find and merge csv files, by defaul are excel files.")
    parser.add_argument("-c", "--merge-cfg", dest="merge_cfg",
                      default=None, help="MergeBOM configuration file.")
    parser.add_argument("-o", "--out-filename", dest="out_filename",
                      default='merged_bom.xlsx', help="Out file name")
    parser.add_argument("-p", "--working-dir", dest="working_dir",
                      default='./', help="BOM to merge working path.")  
    parser.add_argument('--merge_file', '-m', dest='ms', 
                        help='File per il merge', default=None)    
    parser.add_argument('--report_time', '-t', dest='report_time', 
                        help='datetime nel formato : %d/%m/%y', default=None)               
    parser.add_argument("-r", "--bom-revision", dest="bom_rev",
                      default=None, help="Hardware BOM revision")
    parser.add_argument("-pc", "--bom-pcb-revision", dest="bom_pcb_ver",
                      default=None, help="PCB Revision")
    parser.add_argument("-n", "--bom-prj-name", dest="bom_prj_name",
                      default=None, help="Project names.")
    parser.add_argument(
        "-l",
        "--log-on-file",
        dest="log_on_file",
        default=True,
        action="store_true",
        help="List all project name from version file.")
        
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
    options = parser.parse_args()
    options_list['-w', '-a', '-c', '-o', '-p', '-m', '-t', '-r', '-pc', '-n', '-date', '-hw_ver', '-license', '-name', '-name_long', '-pcb', '-pn', '-status',
        options.ws, options.csv_file, options.merge_cfg, options.out_filename, options.working_dir, options.ms, options.report_time, options.bom_rev, options.bom_pcb_ver, options.bom_prj_name, options.log_on_file, options.prj_date, options.prj_hw_ver, options.prj_license, options.prj_name, options.prj_name_long, options.prj_pcb, options.prj_pn, options.prj_status]
    print sys.argv

    f_list= []
    command_line=False
    for i in enumerate(sys.argv):
        for j in enumerate(options_list):
            if not(sys.argv[i]==options_list[j]):
                f_list.append(sys.argv)
                command_line=True
            

    if not command_line:
        if not options.ws == None:
            file_BOM = cfg.cfg_altiumWorkspace(options.ws, options.csv_file) 
            if len(file_BOM)>0:
                appo = []
                for i,v in enumerate(file_BOM):
                    parametri_dict={}
                    appo = file_BOM[i][0]
                    parametri_dict=file_BOM[i][1]
                    options.prj_date=parametri_dict.get('prj_date', None)
                    options.prj_hw_ver=parametri_dict.get('prj_hw_ver',None)
                    options.prj_license=parametri_dict.get('prj_license', None)
                    options.prj_name=parametri_dict.get('prj_name', None)
                    options.prj_name_long=parametri_dict.get('prj_name_long', None)
                    options.prj_pcb=parametri_dict.get('prj_pcb', None)
                    options.prj_pn=parametri_dict.get('prj_pn', None)
                    options.prj_status=parametri_dict.get('prj_status', None)
                    for j,v in enumerate(appo):
                        f_list.append(appo[j])
            else:
                sys.exit
        else:
            print options.ms
            if os.path.exists(options.ms):
                f_list.append(options.ms)
                print f_list
            else:
                sys.exit

    config = cfg.CfgMergeBom(options.merge_cfg)
    if options.report_time is not None:
        options.report_time = datetime.strptime(options.report_time, '%d/%m/%Y')
    
    
    logger = report.Report(log_on_file=options.log_on_file, terminal=True, report_date=options.report_time)
    logger.write_logo()

    m = MergeBom(f_list, config, logger=logger)
    d = m.merge()
    file_list = map(os.path.basename, f_list)
    ft = os.path.join(options.working_dir, options.out_filename)
    report.write_xls(
                    d,
                    file_list,
                    config,
                    ft,
                    hw_ver=options.prj_hw_ver,
                    name=options.prj_name,
                    pcb=options.prj_pcb)

