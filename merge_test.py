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

import unittest
import sys
from merge_xls_bom import *

class TestMergeBom(unittest.TestCase):
    def __init__(self, testname):
        super(TestMergeBom, self).__init__(testname)

    def setUp(self):
        pass
    def tearDown(self):
        pass

    def test_merge(self):
        print "merge"
        file_list = [
            "test/bom_uno.xls",
            "test/bom_due.xls",
            "test/bom_tre.xls",
            "test/bom_quattro.xls",
            "test/test.xlsx"
        ]


        d = import_data(file_list)
        d = group_items(d)
        d = grouped_count(d)
        #header, data = parse_data(file_list)
        #file_list = map(os.path.basename, file_list)
        #stats = write_xls(header, data, file_list, "/tmp/uno.xls")
        for i in d.keys():
            for j in d[i]:
                print j

if __name__ == "__main__":
    #if len(sys.argv) < 2:
    #    printfln("%s <ip addr>" % sys.argv[0])
    #    sys.exit(1)

    suite = unittest.TestSuite()
    suite.addTest(TestMergeBom("test_merge"))
    unittest.TextTestRunner(stream=sys.stdout, verbosity=2).run(suite)

