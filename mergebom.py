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
from lib import cfg,lib,report
from mergebom_class import *
from datetime import datetime
   
if __name__ == "__main__":
    rep = report.Report(log_on_file=True, terminal=True, report_date=None)

    file_list = []
    parser = argparse.ArgumentParser()
    parser.add_argument('--path-workspace-file', '-w', dest='ws', 
                        help='Dove si trova il file WorkSpace', default=None)
<<<<<<< HEAD
    parser.add_argument('--nome-file', '-namef', dest='namef', 
                        help='Nome file da mergiare', default='bom-')
    parser.add_argument('--nome-directory-final-file', '-finalf', dest='finalf', action="store_true",
=======
    parser.add_argument('--nome-file.xlsx', '-nw', dest='nw', 
                        help='Nome file da mergiare', default='bom-')
    parser.add_argument('--caratteristiche-file', '-cf', dest='cf', action="store_true",
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.
                        help='Se il file deve avere lo stesso nome e la stessa directory del file vecchio', default=False)
    parser.add_argument("-a", "--csv", dest="csv_file", action="store_true",
                      default=False, help="Find and merge csv files, by defaul are excel files.")
    parser.add_argument("-c", "--merge-cfg", dest="merge_cfg",
                      default=None, help="MergeBOM configuration file.")
    parser.add_argument("-o", "--out-filename", dest="out_filename",
                      default='merged_bom', help="Out file name")
    parser.add_argument("-p", "--working-dir", dest="working_dir",
                      default="./", help="BOM to merge working path.")     
    parser.add_argument('--report_time', '-t', dest='report_time', 
                        help='datetime nel formato : %d/%m/%y', default=None)               
    parser.add_argument("-r", "--bom-revision", dest="bom_rev",
                      default=None, help="Hardware BOM revision")
    parser.add_argument("-pc", "--bom-pcb-revision", dest="bom_pcb_ver",
                      default=None, help="PCB Revision")
    parser.add_argument("-n", "--bom-prj-name", dest="bom_prj_name",
                      default=None, help="Project names.")
    parser.add_argument("-d", "--delete-file", dest="delete",action="store_true",
                      default=False, help="delete file")                  
<<<<<<< HEAD
<<<<<<< HEAD
=======
>>>>>>> Import diff feature and its test.
    parser.add_argument("-l","--log-on-file",dest="log_on_file",
                      default=True,action="store_true",help="List all project name from version file.")
    parser.add_argument( "-diff","--diff",dest="diff",action="store_true",
                      default=False, help="Generate diff from two specified BOMs")

<<<<<<< HEAD
=======
    parser.add_argument(
        "-l",
        "--log-on-file",
        dest="log_on_file",
        default=True,
        action="store_true",
        help="List all project name from version file.")
        
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.
=======
>>>>>>> Import diff feature and its test.
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

    parser.add_argument('revs', metavar='N', nargs='*', help='revisions', default=None)
    options = parser.parse_args()
<<<<<<< HEAD

<<<<<<< HEAD
    if len(sys.argv) == 1:
        parser.print_help
    if options.cf:
        options.out_filename=options.namef
=======
=======
    
    if len(sys.argv) == 1:
        parser.print_help
>>>>>>> Print short Help when run script without args.
    if options.cf:
        options.out_filename=options.nw
    if len(sys.argv) == 1:
        parser = argparse.ArgumentParser()
        parser.print_help()
<<<<<<< HEAD
        
    
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.
=======
>>>>>>> Add log where needed

    f_list = []
    if options.revs is None or options.revs == []:
        if not options.ws == None:
<<<<<<< HEAD
            file_BOM = cfg.cfg_altiumWorkspace(options.ws, options.csv_file, options.namef, rep)
=======
            file_BOM = cfg.cfg_altiumWorkspace(options.ws, options.csv_file, options.nw)
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.
            print file_BOM
            if len(file_BOM) < 1:
                sys.exit(1)
            appo = []
            for i,v in enumerate(file_BOM):
                parametri_dict={}
                appo = file_BOM[i][0]
                parametri_dict = file_BOM[i][1]
                options.prj_date = parametri_dict.get('prj_date', None)
                options.prj_hw_ver = parametri_dict.get('prj_hw_ver',None)
                options.prj_license = parametri_dict.get('prj_license', None)
                options.prj_name = parametri_dict.get('prj_name', None)
                options.prj_name_long = parametri_dict.get('prj_name_long', None)
                options.prj_pcb = parametri_dict.get('prj_pcb', None)
                options.prj_pn = parametri_dict.get('prj_pn', None)
                options.prj_status = parametri_dict.get('prj_status', None)
                for j,v in enumerate(appo):
                    f_list.append(appo[j])
                
        else:
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
            rep.error("Non è stato trovato nessun file.xlsx o file.csv",
                                self.handler, terminal=self.terminal)
=======
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.
            sys.exit(1)
    else:
        rep.info("Ricerca file mergebom richiesti")
=======
            lib.error("Non è stato trovato nessun file.xlsx o file.csv",
                                self.handler, terminal=self.terminal)
            sys.exit(1)
    else:
        lib.info("Ricerca file mergebom richiesti",
                                self.handler, terminal=self.terminal)
>>>>>>> Add log where needed
=======
            rep.error("Non è stato trovato nessun file.xlsx o file.csv",
                                self.handler, terminal=self.terminal)
            sys.exit(1)
    else:
        rep.info("Ricerca file mergebom richiesti")
>>>>>>> Import diff feature and its test.
        for i,v in enumerate(options.revs):
            f_list.append(options.revs[i])

<<<<<<< HEAD
    if len(f_list) == 0:
        rep.error("Non è stato trovato nessun file da mergiare")
        sys.exit(1)

    if not options.delete:
        if options.prj_hw_ver is None:
=======
    if not options.delete:
<<<<<<< HEAD
        if options.hw_ver is None:
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.
=======
        if options.prj_hw_ver is None:
>>>>>>> Print short Help when run script without args.
            options.out_filename=options.out_filename+'_merge'
        else:
            options.out_filename=options.out_filename+options.prj_hw_ver

<<<<<<< HEAD
    if options.finalf:
=======
    if options.cf:
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.
        appo = f_list[0]
        appo = appo.split(os.sep)
        options.working_dir = os.path.join(*appo[:-1])
    

    config = cfg.CfgMergeBom(options.merge_cfg)
    if options.report_time is not None:
        options.report_time = datetime.strptime(options.report_time, '%d/%m/%Y')
    
<<<<<<< HEAD
<<<<<<< HEAD
 
=======
    lib.info("Inizio operazione di merge",
                                self.handler, terminal=self.terminal)
>>>>>>> Add log where needed
=======
    
>>>>>>> Import diff feature and its test.
     
    logger = report.Report(log_on_file = options.log_on_file, terminal = True, report_date = options.report_time)
    logger.write_logo()

    m = MergeBom(f_list, config, logger=logger)
    d = m.merge()
    file_list = map(os.path.basename, f_list)
    ft = os.path.join(options.working_dir, options.out_filename+'.xlsx')
<<<<<<< HEAD
<<<<<<< HEAD
=======
>>>>>>> Import diff feature and its test.

    if options.diff:
        if len(f_list) != 2:
            logger.error("E' possibile usare la modalità diff solo con 2 file.")
            sys.exit(1)
        d = m.diff()
        l = m.extra_data()
        report.write_xls(
            d,
            file_list,
            config,
            ft,
            diff=True,
            extra_data=l,
            headers=m.header_data())      
    else:    
        rep.info("Inizio operazione di merge")
        report.write_xls(
                        d,
                        file_list,
                        config,
                        ft,
                        hw_ver=options.prj_hw_ver,
                        name=options.prj_name,
                        pcb=options.prj_pcb)

<<<<<<< HEAD
    if options.delete:
        rep.info("Cancellazione vecchio file")
        for i,v in enumerate(f_list):
            os.remove(f_list[i])


=======
    report.write_xls(
                    d,
                    file_list,
                    config,
                    ft,
                    hw_ver=options.prj_hw_ver,
                    name=options.prj_name,
                    pcb=options.prj_pcb)
=======
>>>>>>> Import diff feature and its test.
    if options.delete:
        rep.info("Cancellazione vecchio file")
        for i,v in enumerate(f_list):
            os.remove(f_list[i])
>>>>>>> tutto meno Print short Help when run script without args. Add log where needed.Use pyinstaller to generate binary.Import diff feature and its test.



