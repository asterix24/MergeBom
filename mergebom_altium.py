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
from mergebom import *

try:
    bom_file = sys.argv[1]
    print sys.argv
except IndexError:
    sys.exit(0)

if bom_file:
    bom_file_name = os.path.basename(bom_file)
    bom_file_name = bom_file_name.replace('.xls', '')
    bom_file_name = bom_file_name.replace('.xlsx', '')

    curr_path = os.path.dirname(bom_file)

    report_file = os.path.join(curr_path, bom_file_name + "_report.txt")

    if not bom_file:
        report_file = os.path.join(curr_path, "mergebom_report.txt")

    print "Curr path:", curr_path
    print "Bom file:", bom_file
    print "report file:", report_file

    with open(report_file, 'w') as f:
        f.write(logo_simple)
        f.write("\n")
        f.write("Report file.\n")
        f.write("Date: %s\n" % datetime.datetime.today().strftime("%a %d %B %Y %X"))
        f.write("Directory: %s\n" % curr_path)
        f.write("Revision: %s\n" % "NN")
        f.write("Project Name: %s\n" % "NN")
        f.write("\n")

        if not bom_file:
            f.write("No file found..\n")
            f.close()
            sys.exit(0)

        f.write("%s\n" % bom_file)
        src_bom_file_name = None
        out_bom_file_name = None
        path = os.path.dirname(bom_file)
        name = os.path.basename(bom_file)
        src_bom_file_name = os.path.join(path, "tmp_" + name)
        out_bom_file_name = os.path.join(path, name)

        #print "rename file %s" % out_bom_file_name

        os.rename(out_bom_file_name, src_bom_file_name)
        f.write("SRC file: %s\n" % src_bom_file_name)
        f.write("OUT file: %s\n" % out_bom_file_name)

        f.write("\nCheck Merged items:\n")
        f.write("-" * 80)
        f.write("\n")
        m = MergeBom([src_bom_file_name], handler=f, terminal=False)
        d = m.merge()
        stats = m.statistics()

        write_xls(d, [os.path.basename(src_bom_file_name)],
                  out_bom_file_name, revision="Test", project="Prova")

        f.write("\n\n")
        f.write("=" * 80)
        f.write("\n")

        warning("File num: %s\n" % stats['file_num'], f, terminal=False, prefix="")
        for i in stats.keys():
            if i in CATEGORY_NAMES:
                info(CATEGORY_NAMES[i], f, terminal=False, prefix="- ")
                info("%5.5s %5.5s" % (i, stats[i]), f, terminal=False, prefix="  ")

        f.write("\n")
        f.write("~" * 80)
        f.write("\n")
        warning("Total: %s" % stats['total'], f, terminal=False)

        f.flush()
        f.close()
        # Remove old src file.
        os.remove(src_bom_file_name)

