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

    def test_led(self):
        file_list = [
            "test/bomled.xls",
        ]

        check = {
            'J': [
                [1, 1, u'J1', 'Connector', u'HEADER_2X8_2.54MM_15MM-Stacked_THD', u'Socket Header, 8 pin, 4x2, 2.54mm, H=8.5mm']
            ],
            'D': [
                [3, 3, u'D4, D2, D3', u'BAS70-05', u'SOT-23', u'Diode Dual Schottky Barrier'],
                [9, 9, u'D14, D10, D11, D12, D13, D8, D9, D6, D7', u'+3.3V', u'0603_[1608]_LED', u'Diode LED Green'],
                [2, 2, u'D16, D17', u'S2B', u'DO214AA_12', u'Diode Single'],
                [1, 1, u'D15', u'+2.0V', u'0603_[1608]_LED', u'Diode LED Red'],
                [2, 2, u'D5, D1', u'BAV99', u'SOT-23', u'Diode Dual'],
            ],
            'DZ': [
                [1, 1, u'DZ1', u'B340A', u'DO214AA_12', u'Diode Schottky (STPS2L40U)'],
            ],
          }


        h, d = import_data(file_list)
        d = group_items(d)
        d = grouped_count(d)

        file_list = map(os.path.basename, file_list)
        stats = write_xls(h, d, file_list, "/tmp/uno.xls")

        for category in d.keys():
            for n, i in enumerate(d[category]):
                self.assertEqual(i, check[category][n])

            #for m, x in enumerate(d[i]):
            #    for n, j in enumerate(x):
            #        print j, check[m][n]
            #self.assertEqual(j, check[m][n])

    def test_group(self):
        file_list = [
            "test/bom0.xlsx",
            "test/bom1.xlsx",
            "test/bom2.xlsx",
        ]

        check = [
          [6, 2, 2, 2, u'C101, C201, C200, C1, C0, C100', u'33pF',  u'0603_[1608]', u'Ceramic 50V NP0/C0G'],
          [2, 2, 0, 0, u'C6, C5', u'2.2uF', u'0603_[1608]', u'Ceramic X7R 10V'],
          [3, 3, 0, 0, u'C3, C2, C4', u'1uF', u'1206_[3216]', u'Ceramic X5R 35V, 50V'],
          [4, 0, 2, 2, u'C202, C103, C203, C102', u'100nF', u'0603_[1608]', u'Ceramic X7R 10V'],
        ]


        h, d = import_data(file_list)
        d = group_items(d)
        d = grouped_count(d)

        file_list = map(os.path.basename, file_list)
        stats = write_xls(h, d, file_list, "/tmp/uno.xls")

        for i in d.keys():
            for n, j in enumerate(d[i]):
                print j, check[n]
                for m, c in enumerate(check[n]):
                    print c, j[m]
                    self.assertEqual(c, j[m])

    def test_diff(self):
        file_list = [
            "test/bomdiff1.xlsx",
            "test/bomdiff2.xlsx",
        ]

        h, d = import_data(file_list)
        d = group_items(d)
        a, b = diff_table(d)

        for i in a.keys():
            print i
            for j in a[i]:
                print j

        for i in b.keys():
            print i
            for j in b[i]:
                print j

    def test_orderRef(self):
        test = [
            "C0, C103, C1, C3001, C12",
            "TR100, TR1, TR0, TR10, TR10",
            "RES100, RES1, RES0, RES1001, RES10",
            "SW1",
            "RN1A, RN0B, RN1000, RN3",
        ]
        check = [
            "C0, C1, C12, C103, C3001",
            "TR0, TR1, TR10, TR10, TR100",
            "RES0, RES1, RES10, RES100, RES1001",
            "SW1",
            "RN0B, RN1A, RN3, RN1000",
        ]

        for n,i in enumerate(test):
            l = order_designator(i)
            self.assertEqual(l, check[n])

if __name__ == "__main__":
    suite = unittest.TestSuite()
    #suite.addTest(TestMergeBom("test_group"))
    #suite.addTest(TestMergeBom("test_led"))
    #suite.addTest(TestMergeBom("test_diff"))
    suite.addTest(TestMergeBom("test_orderRef"))
    unittest.TextTestRunner(stream=sys.stdout, verbosity=2).run(suite)

