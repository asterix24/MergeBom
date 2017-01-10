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
import datetime
import ConfigParser
import tempfile
import shutil

from mergebom import *

TEMPLATE = """[<prj>]
name=<name>
hw_ver=<hwver>
pcb_ver=<pcbver>
date=<date>

"""
MERGEBOM_ALTIUM_VER="1.0.0"

class Report(object):
    def __init__(self, directory, logo=None):

        self.src_bom = None

        self.report_file = os.path.join(directory, "mergebom_report.txt")
        self.f = open(self.report_file,'w+')

        if logo is not None:
            self.f.write(logo)
            self.f.write("\n")

        self.f.write("Report file.\n")
	self.f.write("MergeBom Version: %s\n" % MERGEBOM_VER)
	self.f.write("MergeBom Altium Version: %s\n" % MERGEBOM_ALTIUM_VER)
        dt = datetime.datetime.now()
        self.f.write("Date: %s\n" % dt.strftime("%A, %d %B %Y %X"))
        self.f.write("." * 80)
        self.f.write("\n")

    def __del__(self):
        self.f.flush()
        self.f.close()

    def write_header(self, d, file_list):
        self.f.write("\n")
        self.f.write(":" * 80)
        self.f.write("Date: %s\n" %              d['date'])
        self.f.write("Project Name: %s\n" %      d['name'])
        self.f.write("Hardware Revision: %s\n" % d['hw_ver'])
        self.f.write("PCB Revision: %s\n" %      d['pcb_ver'])
        self.f.write("\n")

        self.f.write("Bom Files:\n")
        for i in file_list:
            self.f.write(" - %s\n" % i)

        self.f.write("\n== Check Merged items: ==\n")
        self.f.write("-" * 80)
        self.f.write("\n")

    def write_stats(self, stats):
        self.f.write("\n\n")
        self.f.write("=" * 80)
        self.f.write("\n")

        self.f.write("File num: %s\n" % stats['file_num'])
        for i in stats.keys():
            if i in CATEGORY_NAMES:
                self.f.write(CATEGORY_NAMES[i] + "\n")
                self.f.write("%5.5s %5.5s\n" % (i, stats[i]))

        self.f.write("\n")
        self.f.write("~" * 80)
        self.f.write("\n")
        self.f.write("Total: %s\n" % stats['total'])

    def error(self, msg):
        self.f.write("Error:")
        self.f.write("%s", msg)
        self.f.write("\n")

    def handler(self):
        return self.f


def read_ini(config, section_name):
    d = {}
    d['name' ] = config.get(section_name, 'name')
    d['hw_ver'   ] = config.get(section_name, 'hw_ver')
    d['pcb_ver'  ] = config.get(section_name, 'pcb_ver')
    d['date'     ] = config.get(section_name, 'date')

    return d

if __name__ == "__main__":
    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option("-i", "--init", dest="init_flag", action="store_true", help="Create version file at first time.")
    parser.add_option("-n", "--prj-name", dest="prj_name", default='MyProject', help="Project names.")
    parser.add_option("-b", "--bom-file", dest="bom_file", default="bom.xlsx", help="BOM file name.")
    parser.add_option("-v", "--version-file", dest="version_file", default='version.txt', help="Version file.")
    parser.add_option("-d", "--search-dir", dest="directory", default='./', help="Directory path to seach.")
    parser.add_option("-o", "--out-bomname", dest="out_bom_file", default=None, help="Directory path to seach.")
    (options, args) = parser.parse_args()

    print args, options

    if options.init_flag:
        with open(os.path.join(options.directory, options.version_file), "w") as f:
            vv = TEMPLATE.replace("<prj>", options.prj_name)
            vv = vv.replace("<name>", options.prj_name)
            vv = vv.replace("<date>", datetime.datetime.today().strftime("%Y%m%d"))
            f.write(vv)
            f.flush()
        print "File correctly generated."
        sys.exit(0)

    # generate report file, and init it
    report = Report(options.directory, logo=logo_simple)

    # search version file
    file_list = glob.glob(os.path.join(options.directory, options.version_file))
    if not file_list and len(file_list) != 1:
        print "Error No version file found!"
        report.error("Error: No version file found!\n")
        sys.exit(1)

    # Whit version file search all bom files
    version_file = file_list[0]
    config = ConfigParser.ConfigParser()
    config.readfp(open(version_file))

    for section in  config.sections():
        ini_data = read_ini(config, section)

        # Search bom file, in section we fount path
        for search_path in [ (options.directory,"Assembly"),
                    (options.directory,"Assembly", section) ]:

            # current path
            path = os.path.join(*search_path)
            file_list = glob.glob(os.path.join(path, '*.xls'))
            file_list += glob.glob(os.path.join(path, '*.xlsx'))

            if not file_list:
                continue

            report.write_header(ini_data, file_list)
            tmpdir = tempfile.gettempdir()

            process_file = []
            for i in file_list:
                name = os.path.basename(i)
                dst = os.path.join(tmpdir, name)
                shutil.copy(i, dst)
                process_file.append(dst)
                os.remove(i)
                print dst

            m = MergeBom(process_file, handler=report.handler(), terminal=False)
            d = m.merge()

            stats = m.statistics()
            report.write_stats(stats)
            st = []
            for i in stats.keys():
                if i in CATEGORY_NAMES:
                    st.append((stats[i], CATEGORY_NAMES[i]))
            st.append((stats['total'], "Total"))

            print "section", section
            bom_file_name = os.path.join(path, "bom-%s.xlsx" % section)
            print "bom file", bom_file_name
            if options.out_bom_file is not None:
                bom_file_name = options.out_bom_file

            write_xls(d, map(os.path.basename, file_list), bom_file_name,
                      hw_ver=ini_data['hw_ver'], pcb_ver=ini_data['pcb_ver'],
                      project=ini_data['name'],
                      statistics=st)

            # Remove old src file.
            for i in process_file:
                os.remove(i)


