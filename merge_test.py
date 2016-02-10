#! /usr/bin/env python
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

import unittest
import sys
from mergebom import *


class TestMergeBom(unittest.TestCase):
    def __init__(self, testname):
        super(TestMergeBom, self).__init__(testname)

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_import(self):
        file_list = [
            "test/bom0.xlsx",
            "test/bom1.xlsx",
            "test/bom2.xlsx",
        ]
        check = [
            (7,),
            (4,),
            (4,),
        ]

        l = MergeBom(file_list)
        data = l.table_data()
        self.assertEqual(len(data), 3)
        print
        for n, i in enumerate(data):
            self.assertEqual(len(i), check[n][0])
            print len(i)

    def test_led(self):
        file_list = [
            "test/bomled.xls",
        ]

        check = {
            'J': [
                [1, 1, u'J1', 'Connector', u'HEADER_2X8_2.54MM_15MM-Stacked_THD', u'Socket Header, 8 pin, 4x2, 2.54mm, H=8.5mm']
            ],
            'D': [
                [3, 3, u'D2, D3, D4', u'BAS70-05', u'SOT-23', u'Diode Dual Schottky Barrier'],
                [9, 9, u'D6, D7, D8, D9, D10, D11, D12, D13, D14', u'+3.3V', u'0603_[1608]_LED', u'Diode LED Green'],
                [2, 2, u'D16, D17', u'S2B', u'DO214AA_12', u'Diode Single'],
                [1, 1, u'D15', u'+2.0V', u'0603_[1608]_LED', u'Diode LED Red'],
                [2, 2, u'D1, D5', u'BAV99', u'SOT-23', u'Diode Dual'],
            ],
            'DZ': [
                [1, 1, u'DZ1', u'B340A', u'DO214AA_12', u'Diode Schottky (STPS2L40U)'],
            ],
          }

        m = MergeBom(file_list)
        d = m.merge()

        # file_list = map(os.path.basename, file_list)
        # stats = write_xls(h, d, file_list, "/tmp/uno.xls")

        for category in d.keys():
            for n, i in enumerate(d[category]):
                print "T >", i
                print "C <", check[category][n]
                self.assertEqual(i, check[category][n])

    def test_group(self):
        file_list = [
            "test/bom0.xlsx",
            "test/bom1.xlsx",
            "test/bom2.xlsx",
        ]

        check = [
          [6, 2, 2, 2, u'C0, C1, C100, C101, C200, C201', u'33pF',  u'0603_[1608]', u'Ceramic 50V NP0/C0G'],
          [4, 0, 2, 2, u'C102, C103, C202, C203', u'100nF', u'0603_[1608]', u'Ceramic X7R 10V'],
          [3, 3, 0, 0, u'C2, C3, C4', u'1uF', u'1206_[3216]', u'Ceramic X5R 35V, 50V'],
          [2, 2, 0, 0, u'C5, C6', u'2.2uF', u'0603_[1608]', u'Ceramic X7R 10V'],
        ]

        m = MergeBom(file_list)
        d = m.merge()

        # file_list = map(os.path.basename, file_list)
        # stats = write_xls(h, d, file_list, "/tmp/uno.xls")

        for i in d.keys():
            for n, j in enumerate(d[i]):
                print "T >", j
                print "C <", check[n]
                for m, c in enumerate(check[n]):
                    print "T >", c
                    print "C <", j[m]
                    self.assertEqual(c, j[m])

    def test_diff(self):
        file_list = [
            "test/bomdiff1.xlsx",
            "test/bomdiff2.xlsx",
        ]

        check = {
            'C45': (
              ['bomdiff1.xlsx', '-', '-', '-', '-', '-'],
              ['bomdiff2.xlsx', 1, u'C45', u'Ceramic 50V NP0/C0G', u'1nF', u'0603_[1608]'],
            ),
            'C1045': (
              ['bomdiff1.xlsx', 1, u'C1045', u'Ceramic 50V NP0/C0G', u'100nF', u'0603_[1608]'],
              ['bomdiff2.xlsx', '-', '-', '-', '-', '-'],
            ),
            'C204': (
              ['bomdiff1.xlsx', 1, u'C204', u'Ceramic 50V NP0/C0G', u'18pF', u'0603_[1608]'],
              ['bomdiff2.xlsx', '-', '-', '-', '-', '-'],
            ),
            'C2046': (
              ['bomdiff1.xlsx', '-', '-', '-', '-', '-'],
              ['bomdiff2.xlsx', 1, u'C2046', u'Ceramic 50V NP0/C0G', u'18pF', u'0603_[1608]'],
            ),
            'C104': (
              ['bomdiff1.xlsx', 1, u'C104', u'Ceramic 50V NP0/C0G', u'100nF', u'0603_[1608]'],
              ['bomdiff2.xlsx', 1, u'C104', u'Ceramic 50V NP0/C0G', u'10nF', u'0603_[1608]'],
            ),
            'C1': (
              ['bomdiff1.xlsx', '-', '-', '-', '-', '-'],
              ['bomdiff2.xlsx', 1, u'C1', u'Tantalum 10V Low ESR (TPSP106M010R2000)', u'10uF Tantalum', u'0805_[2012]_POL'],
            )
            }

        m = MergeBom(file_list)
        k = m.diff()

        print
        for i in k.keys():
            print i, ">>", k[i][0]
            print i, "<<", k[i][1]
            print "~" * 80
            self.assertEqual(k[i][0], check[i][0])
            self.assertEqual(k[i][1], check[i][1])

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

        for n, i in enumerate(test):
            l = order_designator(i)
            self.assertEqual(l, check[n])

    def test_valueToFloat(self):
        test = [
            ("1R1","1R2", "0R3", "1R8"),
            ("1k", "1k5", "1", "10R", "1R2", "2.2k", "0.3"),
            ("0.1uF", "100nF", "1F", "10pF", "2.2uF", "47uF", "1uF"),
            ("1nH", "1H", "10pH", "2.2uH", "47uH"),
        ]
        checkv1 = [
            (0.3, 1.1, 1.2, 1.8),
            (0.3, 1, 1.2, 10, 1000, 1500, 2200),
            (10e-12, 100e-9, 0.1e-6, 1e-6, 2.2e-6, 47e-6, 1),
            (10e-12, 1e-9, 2.2e-6, 47e-6, 1),
        ]

        for k, m in enumerate(test):
            l = order_value(m)
            for n, i in enumerate(l):
                print i, "->", checkv1[k][n]
                self.assertTrue(abs((i - checkv1[k][n]) < 0.1))
                print "-" * 80

    def test_floatToValue(self):
        test = [
            ((1000, ""), (1500, ""), (2200, ""), (2210, ""), (4700, ""), (47000, "")),
            ((1000000, ""), (1500000, ""), (860000, ""), (8600000, "")),
            ((1.2, ""), (3.33, ""), (0.12, ""), (1.234, ""), (0.33, ""), (3.3e-8,"")),
            ((0.000010, "F"), (1e-3, "F"), (10e-3,"H")),
            ((1.5e-6, "F"), (33e-12, "F"), (100e-9, "F")),
            ((1, ""), (0.1, "")),
        ]
        check = [
            ("1k","1k5", "2k2", "2k21", "4k7", "47k"),
            ("1M", "1M5", "860k", "8M6"),
            ("1R2", "3R33", "120m", "1R234", "330m", "33n"),
            ("10uF", "1mF", "10mH"),
            ("1.5uF", "33pF", "100nF"),
            ("1R", "100m"),
        ]

        for k, m in enumerate(test):
            l = value_toStr(m)
            self.assertTrue(l)
            for n, i in enumerate(l):
                print i, "->", check[k][n]
                self.assertEqual(i, check[k][n])
                print "-" * 80

    def test_outFile(self):
        file_list = [
            "test/bom_uno.xls",
            "test/bom_due.xls",
            "test/bom_tre.xls",
            "test/bom_quattro.xls",
        ]

        m = MergeBom(file_list)
        d = m.merge()
        file_list = map(os.path.basename, file_list)
        write_xls(d, file_list, "/tmp/uno.xls", hw_ver="13", pcb_ver="C", project="TEST")

    def test_mergedFile(self):
        file_list = [
            "test/bom-merged.xls",
        ]

        m = MergeBom(file_list)
        d = m.extra_data()
        print d
        self.assertEqual(len(d), 1)
        self.assertEqual(d[0]['project'], "test")
        self.assertEqual(d[0]['pcb_version'], "c")

    def test_stats(self):
        file_list = [
            "test/bom-merged.xls",
        ]

        info(logo, sys.stdout, terminal=True, prefix="")
        m = MergeBom(file_list)
        m.merge()
        stats = m.statistics()
        warning("File num: %s" % stats['file_num'], sys.stdout, terminal=True)
        for i in stats.keys():
            if i in CATEGORY_NAMES:
                info(CATEGORY_NAMES[i], sys.stdout, terminal=True, prefix="- ")
                info("%5.5s %5.5s" % (i, stats[i]), sys.stdout, terminal=True, prefix="  ")

        warning("Total: %s" % stats['total'], sys.stdout, terminal=True)

    def test_notPopulate(self):
        """
        Cerca i componenti da non montare
        """
        file_list = [
            "test/bom-np.xls",
        ]

        m = MergeBom(file_list)
        d = m.merge()
        stats = m.statistics()
        st = []
        for i in stats.keys():
            if i in CATEGORY_NAMES:
                st.append((stats[i], CATEGORY_NAMES[i]))
        st.append((stats['total'], "Total"))

        file_list = map(os.path.basename, file_list)
        write_xls(d, file_list, "/tmp/uno.xlsx", hw_ver="13",
                  pcb_ver="C", project="TEST", statistics=st)

    def test_version(self):
        import subprocess
        import mergebom_altium
        import ConfigParser

        subprocess.call(["python", "./mergebom_altium.py", "-i",
                         "-v", "prova.txt", "-d", "/tmp"])

        check = {
            'name': "MyProject",
            'hw_ver': "<hwver>",
            'pcb_ver': "<pcbver>",
            'date': "",
        }

        config = ConfigParser.ConfigParser()
        config.readfp(open("/tmp/prova.txt"))
        d = mergebom_altium.read_ini(config, "MyProject")

        self.assertTrue("date" in d)
        for i in ["name", "hw_ver", "pcb_ver"]:
            self.assertEqual(d[i], check[i])

if __name__ == "__main__":
    suite = unittest.TestSuite()
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_group"))
    suite.addTest(TestMergeBom("test_led"))
    suite.addTest(TestMergeBom("test_diff"))
    suite.addTest(TestMergeBom("test_orderRef"))
    suite.addTest(TestMergeBom("test_outFile"))
    suite.addTest(TestMergeBom("test_mergedFile"))
    suite.addTest(TestMergeBom("test_stats"))
    suite.addTest(TestMergeBom("test_valueToFloat"))
    suite.addTest(TestMergeBom("test_floatToValue"))
    suite.addTest(TestMergeBom("test_version"))
    suite.addTest(TestMergeBom("test_notPopulate"))
    unittest.TextTestRunner(stream=sys.stdout, verbosity=2).run(suite)
