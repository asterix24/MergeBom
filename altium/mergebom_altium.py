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
#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys, os
import glob

curr_path = None
for n, i in enumerate(sys.argv):
    if n:
        curr_path = os.path.dirname(i)

if curr_path is not None:
    report_file = "mergebom_report.txt"
    with open(report_file, 'w') as f:
        f.write("Argomenti: %s\r\n" % len(sys.argv))
        f.write("Directory: %s\r\n" %  curr_path)

        if curr_path is not None:
            file_list = glob.glob(os.path.join(curr_path, '*.xls'))
            file_list += glob.glob(os.path.join(curr_path, '*.xlsx'))

            for i in file_list:
                f.write("file: %s\r\n" % i)
