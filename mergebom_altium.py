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
from mergebom import *

curr_path = None
for n, i in enumerate(sys.argv):
    if n:
        curr_path = i

if curr_path is not None:
    report_file = os.path.join(curr_path, "mergebom_report.txt")
    print report_file

    with open(report_file, 'w') as f:
        f.write("Argomenti: %s\n" % len(sys.argv))
        f.write("Directory: %s\n" % curr_path)
        f.write("\n")

        file_list = glob.glob(os.path.join(curr_path, '*.xls'))
        file_list += glob.glob(os.path.join(curr_path, '*.xlsx'))

        print file_list, len(file_list)
        assert(len(file_list) == 1)

        src_bom_file_name = None
        out_bom_file_name = None
        for i in file_list:
            path = os.path.dirname(i)
            name = os.path.basename(i)
            src_bom_file_name = os.path.join(path, "tmp_" + name)
            out_bom_file_name = os.path.join(path, name)

            print "rename file %s" % out_bom_file_name

            os.rename(out_bom_file_name, src_bom_file_name)
            f.write("SRC file: %s\n" % src_bom_file_name)
            f.write("OUT file: %s\n" % out_bom_file_name)

        m = MergeBom([src_bom_file_name], handler=f, terminal=False)
        d = m.merge()
        m.statistics()

        write_xls(d, [os.path.basename(src_bom_file_name)],
                  out_bom_file_name, revision="Test", project="Prova")

        # Remove old src file.
        os.remove(src_bom_file_name)
