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

import sys
import os
import unittest
import subprocess
import tempfile
import argparse
import glob

from lib import cfg, report, lib
from mergebom_class import *
import tempfile
from datetime import datetime

def dump(d):
    for i in d:
        print("Key: %s" % i)
        print("Rows [%d]:" % len(d[i]))
        for j in d[i]:
            print(j)
        print("-" * 80)

class TestMergeBom(unittest.TestCase):
    """
    MergeBOM Test suite.
    """

    def __init__(self, testname):
        super(TestMergeBom, self).__init__(testname)
        self.logger = None
        self.config = None

    def setUp(self):
        self.logger = report.Report(log_on_file=True, terminal=True, report_date=None)
        self.config = cfg.CfgMergeBom()
        self.temp_dir = tempfile.gettempdir()

    def tearDown(self):
        pass

    def test_altiumWorkspace(self):
        file_BOM = [
            ('progettotest1',[], {
                'prj_status': 'status',
                'prj_pcb': 'C',
                'prj_name': 'TEST',
                'prj_date': '28/05/2018',
                'prj_pn': 'pn',
                'prj_name_long': 'CMOS Sensor adapter iMX8',
                'prj_license': '-',
                'prj_hw_ver': '13'
                }),
            ('progettotest2',[], {
                'prj_status': 'status',
                'prj_pcb': 'A',
                'prj_name': 'Adapter-imx8',
                'prj_date': '28/05/2018',
                'prj_pn': 'pn',
                'prj_name_long': 'CMOS Sensor adapter iMX8',
                'prj_license': '-',
                'prj_hw_ver': '0'
                }
            ),
        ]


        p = os.path.join('test','utils.DsnWrk')
        param = cfg.cfg_altiumWorkspace(p, False, "Assembly", self.logger,
                                        bom_prefix='bom-', bom_postfix="")
        self.assertEqual(file_BOM, param)

        check = [
            ('5mp-sensor', [],
             {'prj_pcb': 'A',
              'prj_name': 'Sensor Shield 5Mp',
              'prj_date': '01/06/2018',
              'prj_status': 'Prototype',
              'prj_name_long': 'Camera sensor shield 5Mp AR0521',
              'prj_license': 'Copyright company spa',
              'prj_hw_ver': '0', 'prj_prefix': ''}),
            ('camera-core', [os.path.join(os.path.abspath(os.path.dirname(__file__)), 'test','wktest','Assembly', 'camera-core', 'bom-camera-core.xlsx')],
             {'prj_status': 'Prototype',
              'prj_pcb': 'A',
              'prj_name': 'Camera Core',
              'prj_date': '01/06/2018',
              'prj_pn': '-',
              'prj_name_long': 'Camera core imx8',
              'prj_license': 'Copyright company spa',
              'prj_hw_ver': '0',
              'prj_prefix': ''}),
            ('18mp-sensor', [],
             {'prj_pcb': 'A',
              'prj_name': 'Sensor Shield 18Mp',
              'prj_date': '01/06/2018',
              'prj_status': 'Prototype',
              'prj_name_long': 'Camera sensor shield 5Mp AR1820',
              'prj_license': 'Copyright company spa',
              'prj_hw_ver': '0',
              'prj_prefix': ''}),
        ]

        wk_path = os.path.abspath(os.path.join('test', "wktest", "camera.DsnWrk"))
        #wk_path = os.path.join('test', "wktest", "camera.DsnWrk")
        param = cfg.cfg_altiumWorkspace(wk_path, False, "Assembly", self.logger,
                                        bom_prefix='bom-', bom_postfix="")
        self.assertEqual(check, param)

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

        l = MergeBom(file_list, self.config, logger=self.logger)
        data = l.table_data()
        self.assertEqual(len(data), 3)
        for n, i in enumerate(data):
            self.assertEqual(len(i), check[n][0])

    def test_led(self):
        file_list = [
            "test/bomled.xls",
        ]

        check = {
            'J': [
                [1, 1, u'J1', 'Connector', u'HEADER_2X8_2.54MM_15MM-Stacked_THD', u'Socket Header, 8 pin, 4x2, 2.54mm, H=8.5mm']
            ],
            'D': [
                [1, 1, u'DZ1', u'B340A', u'DO214AA_12', u'Diode Schottky (STPS2L40U)'],
                [3, 3, u'D2, D3, D4', u'BAS70-05', u'SOT-23', u'Diode Dual Schottky Barrier'],
                [2, 2, u'D1, D5', u'BAV99', u'SOT-23', u'Diode Dual'],
                [1, 1, u'D15', u'LED', u'0603_[1608]_LED', u'Diode LED Red'],
                [9, 9, u'D6, D7, D8, D9, D10, D11, D12, D13, D14', u'LED', u'0603_[1608]_LED', u'Diode LED Green'],
                [2, 2, u'D16, D17', u'S2B', u'DO214AA_12', u'Diode Single'],
            ],
        }

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()

        # file_list = map(os.path.basename, file_list)
        # stats = report.write_xls(h, d, file_list, "/tmp/uno.xls")

        for category in d.keys():
            for n, i in enumerate(d[category]):
                print("T >", i)
                print("C <", check[category][n])
                self.assertEqual(i, check[category][n])

    def test_rele(self):
        file_list = [
            "test/bomrele.xls",
        ]

        check = {
            'J': [
                [1, 1, u'J35', 'Connector', u'DSUB1.385-2H25A_FEMALE', u'Receptacle Assembly, 25 Position, Right Angle'],
                [2, 2, u'J38, J43', 'Connector', u'TBLOCK_1X8_3.5MM_HOR_THD', u'Terminal block, Header, 5.08mm, 8x1, Vertical + Plug'],
                [9, 9, u'J5, J9, J11, J15, J17, J19, J25, J30, J42', 'Connector', u'TBLOCK_1X3_3.5MM_HOR_THD', u'Terminal block, Header, 5.08mm, 3x1, Vertical + Plug'],
                [1, 1, u'J3', 'Connector', u'HEADER_2x10_2.00MM_BOX_LEV_THD', u'Header, 10-Pin, Dual row, 2mm, Boxed with Lever'],
                [3, 3, u'J48, J49, J50', 'Connector', u'SOCKET_2x26_2.54MM_THD', u'Socket Header, 52 pin, 26x2, 2.54mm, H=8.5mm'],
                [1, 1, u'J12', 'Connector', u'HEADER_2x10_2.00MM_BOX_THD', u'Header, 8-Pin, Dual row, 2mm, Boxed'],
                [19, 19, u'J4, J10, J18, J22, J23, J27, J31, J32, J33, J34, J36, J37, J39, J40, J41, J44, J45, J46, J47', 'Connector', u'TBLOCK_1X2_3.5MM_HOR_THD', u'Terminal block, Header, 5.08mm, 2x1, Vertical + Plug'],
                [28, 28, u'J21, J24, J26, J29, J51, J52, J53, J54, J55, J56, J57, J58, J59, J60, J61, J62, J63, J64, J65, J66, J67, J68, J69, J70, J71, J72, J73, J74', 'Connector', u'HEADER_1X3_2.54MM_THD', u'Pin Header, 3x1, 2.54mm, THD'],
                [10, 10, u'J8, J16, J20, J28, J75, J76, J77, J78, J79, J80', 'Connector', u'HEADER_1X2_2.54MM_THD', u'Pin Header, 2x1, 2.54mm, THD'],
            ],
            'U': [
                [1, 1, u'U1', u'24C128', u'SOIC8', u'EEPROM'],
                [8, 8, u'U19, U21, U22, U23, U25, U27, U28, U29', u'4N25', u'SOIC6_OPTO', u'Optocoupler'],
                [4, 4, u'U14, U15, U16, U17', u'74HC4051', u'SOIC16', u'Analog Mux 8:1'],
                [3, 3, u'U11, U20, U26', u'AD5624RBRMZ', u'MSOP50P490X110-10P', u'12bit DAC'],
                [2, 2, u'U31, U33', u'AD799x', u'TSOP65P640X120-20L', u'ADC 10/12bit I2C'],
                [6, 6, u'U32, U34, U35, U36, U37, U47', u'AD8418BRMZ', u'MSOP50P490X110-10P', u'12bit DAC'],
                [1, 1, u'U6', u'ADS7951SBDBT', u'TSSOP50P640-30L', u'ADC'],
                [2, 2, u'U3, U4', u'LM22670TJ-ADJ', u'TO-263-7', u'Switching regulator.'],
                [1, 1, u'U2', u'LM22671MR-ADJ', u'SOIC8_PAD', u'DC-DC switch converter'],
                [3, 3, u'U12, U18, U24', u'OPA4188AID', u'SOIC14', u'Quad Op Amp'],
                [1, 1, u'U13', u'OPA4188AID', u'SOIC14', u'Op Amp'],
                [1, 1, u'U5', u'REF3325AIDBZR', u'SOT-23', u'Shunt volt reference +2.5V'],
                [6, 6, u'U38, U39, U42, U43, U44, U46', u"Relay, Rele'", u'relays_spco_thd_28x12mm_5mm', u'Rele NT75CS16DC12V0.415.0 16A 12V'],
                [1, 1, u'U45', u"Relay, Rele'", u'relays_spco_thd_28x12mm_5mm', u'Rele NT75CS16DC12V0.415.0 16A 24V'],
                [1, 1, u'U41', u"Relay, Rele'", u'relays_spco_thd_28x12mm_5mm', u'Rele Altra marca 16A 12V'],
                [1, 1, u'U30', u'TLV431 SOT23', u'SOT-23', u'TLV431BQDBZT'],
                [1, 1, u'U40', u'ULN2803', u'SOIC18', u'BJT Darlinton Array'],
                [4, 4, u'U7, U8, U9, U10', u'XTR117_DGK', u'MSOP65P490X110-8P', u'4-20mA loop Trasmitter'],
            ],
            'D': [
                [1, 1, u'D8', u'B340A', u'DO214AA_12', u'Diode Schottky (STPS2L40U)'],
                [4, 4, u'D9, D10, D11, D12', u'BAS70-05', u'SOT-23', u'Schottky Barrier Double Diodes'],
                [3, 3, u'D13, D14, D16', u'BAV99', u'SOT-23', u'BAV99 Diode'],
                [2, 2, u'D1, D7', 'LED', u'0603_[1608]_K2-A1_LED', u'LED Green'],
                [6, 6, u'D4, D5, D6, D15, D17, D18', 'LED', u'0603_[1608]_K2-A1_LED', u'Diode LED Green'],
                [1, 1, u'D2', u'S2B', u'DO214AA_12', u'S2B Diode'],
                [1, 1, u'D3', u'SMBJ28A', u'DO214AA_12', u'TRANSIL (SM6T12A)'],
            ],
        }

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()

        # file_list = map(os.path.basename, file_list)
        # stats = report.write_xls(h, d, file_list, "/tmp/uno.xls")

        for category in d.keys():
            for n, i in enumerate(d[category]):
                print("T >", i)
                print("C <", check[category][n])
                self.assertEqual(i, check[category][n])

    def test_group(self):
        file_list = [
            "test/bom0.xlsx",
            "test/bom1.xlsx",
            "test/bom2.xlsx",
        ]

        check = [[6,
                  2,
                  2,
                  2,
                  u'C0, C1, C100, C101, C200, C201',
                  u'33pF',
                  u'0603_[1608]',
                  u'Ceramic 50V NP0/C0G'],
                 [4,
                  0,
                  2,
                  2,
                  u'C102, C103, C202, C203',
                  u'100nF',
                  u'0603_[1608]',
                  u'Ceramic X7R 10V'],
                 [3,
                  3,
                  0,
                  0,
                  u'C2, C3, C4',
                  u'1uF',
                  u'1206_[3216]',
                  u'Ceramic X5R 35V, 50V'],
                 [2,
                  2,
                  0,
                  0,
                  u'C5, C6',
                  u'2.2uF',
                  u'0603_[1608]',
                  u'Ceramic X7R 10V'],
                 ]

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()

        # file_list = map(os.path.basename, file_list)
        # stats = report.write_xls(h, d, file_list, "/tmp/uno.xls")

        for i in d.keys():
            for n, j in enumerate(d[i]):
                print("T >", j)
                print("C <", check[n])
                for m, c in enumerate(check[n]):
                    print("T >", c)
                    print("C <", j[m])
                    self.assertEqual(c, j[m])

    def test_groupFmt(self):
        file_list = [
			os.path.join('test', 'bom-fmt.xls'),
        ]

        check = {
            "10k" : (1, 9),
            "fpf2124": (1, 1),
            "txs0108e": (1, 1),
            "1uf": (1, 3),
            "tactile switch": (1, 1),
            "100k": (1, 7),
            "sp2526a-1en-l": (1, 1),
            "b340a": (1, 1),
            "wl1831mod": (1, 1),
            "bridge diode": (1, 1),
            "smbj28a": (1, 1),
            "mbr120vlsft1g": (1, 2),
            "np  (1k)": (1, 1),
            "68k": (1, 31),
            "np connector": (1, 1),
            "connector": (11, 12),
            "tvs array (we 82400152)": (1, 2),
            "10uf": (1, 10),
            "10nf": (1, 4),
            "np": (1, 15),
            "txb0108pw": (1, 1),
            "np  (irlml6402)": (1, 1),
            "led": (1, 5),
            "develboard": (1, 1),
            "ld29150dt33r": (1, 1),
            "220h r @ 100mhz": (1, 14),
            "100nf": (1, 30),
            "6.8r": (1, 3),
            "ft232rq": (1, 1),
            "line filter (we 744232090)": (1, 2),
            "np  (220 r @ 100mhz)": (1, 1),
            "100uf": (2, 2),
            "lm22670mre-5.0/nopb": (1, 1),
            "0r": (1, 1),
            "4.7uf": (2, 2),
            "np  (100nf)": (1, 1),
            "180h @ 100mhz": (1, 3),
            "470r": (1, 3),
            "lm1117-18": (1, 1),
            "np  (10k)": (1, 1),
            "10uh  we 74437334100": (1, 1),
            "32.768khz oscillator": (1, 1),
        }

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()

        test_dict = {}
        row_count = 0
        count = 0
        for i in d.keys():
            for n, j in enumerate(d[i]):
                key = j[3].lower().strip()
                if key in test_dict:
                    row_count, count = test_dict[key]
                    row_count += 1
                    count += j[0]
                    test_dict[key] = (row_count, count)
                else:
                    test_dict[key] = (1, j[0])

        for k in test_dict.keys():
            print("\"%s\": (%s, %s)," % (k, test_dict[k][0], test_dict[k][1]))
            print("T > %20.20s | row count %5.5s | count %5.5s |" % (k, test_dict[k][0], test_dict[k][1]))
            print("C < %20.20s | row count %5.5s | count %5.5s |" % (k, check[k][0], check[k][1]))

            self.assertEqual(check[k][0], test_dict[k][0], \
                             "Unable to merge row whit differt componet value. [%s]" % k)
            self.assertEqual(check[k][1], test_dict[k][1], \
                             "Unable to merge row whit differt componet value. [%s]" % k)

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

        m = MergeBom(file_list, self.config, logger=self.logger)
        k = m.diff()

        print()
        for i in k.keys():
            print(i, ">>", k[i][0])
            print(i, "<<", k[i][1])
            print("~" * 80)
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
            l = lib.order_designator(i, self.logger)
            self.assertEqual(l, check[n])

    def test_valueToFloat(self):
        test = [
            ("C", ("100pF", )),
            ("R", ("1R1", "1R2", "0R3", "1R8")),
            ("R", ("1.1R", "1.2R", "0.3R", "1.8R")),
            ("R", ("1k", "1k5", "1", "10R", "1R2", "2.2k", "0.3")),
            ("C", ("0.1uF", "100nF", "1F", "10pF", "2.2uF", "47uF", "1uF")),
            ("L", ("1nH", "1H", "10pH", "2.2uH", "47uH")),
            ("R", ("68ohm", "1.2R", "3.33R", "0.12R", "1.234R", "0.33R", "33nohm")),
            ("Y", ("32.768kHz", "1MHz", "12.134kHz")),
            ("L", ("4mH", "4.7mH", "100uH")),
            ("R", ("33nohm",)),
            ("R", ("2k31", "10k12", "5K421", "4R123", "1M12")),
        ]
        checkv1 = [
            ((1e-10, "F"),),
            ((0.3, "ohm"), (1.1, "ohm"), (1.2, "ohm"), (1.8, "ohm")),
            ((0.3, "ohm"), (1.1, "ohm"), (1.2, "ohm"), (1.8, "ohm")),
            ((0.3, "ohm"), (1.0, "ohm"), (1.2, "ohm"), (10.0, "ohm"), (1000.0, "ohm"), (1500.0, "ohm"), (2200.0, "ohm")),
            ((10e-12, "F"), (100e-9, "F"), (0.1e-6, "F"), (1e-6, "F"), (2.2e-6, "F"), (47e-6, "F"), (1.0, "F")),
            ((10e-12, "H"), (1e-9, "H"), (2.2e-6, "H"), (47e-6, "H"), (1.0, "H")),
            ((3.3e-8, "ohm"), (0.12, "ohm"), (0.33, "ohm"), (1.2, "ohm"), (1.234, "ohm"), (3.33, "ohm"), (68, "ohm")),
            ((12134.0, "Hz"), (32768.0, "Hz"), (1e6, "Hz")),
            ((100e-6, "H"), (4e-3, "H"), (0.0047, "H")),
            ((33e-9, "ohm"),),
            ((4.123, "ohm"), (2310.0, "ohm"), (5421.0, "ohm"), (10120.0, "ohm"), (1120000.0, "ohm")),
        ]

        print()
        for k, m in enumerate(test):
            l = []
            for mm in m[1]:
                a, b, _ = lib.value_toFloat(mm, m[0], self.logger)
                l.append((a, b))

            l = sorted(l, key=lambda x: x[0])
            for n, i in enumerate(l):
                print(i, "->", checkv1[k][n])
                self.assertTrue(abs(i[0] - checkv1[k][n][0]) < 10e-9)
                self.assertEqual(i[1], checkv1[k][n][1])
                print("-" * 80)

    def test_floatToValue(self):
        test = [
            ((1000, "ohm", ""), (1500, "ohm", ""), (2200, "ohm", ""), (2210, "ohm", ""), (4700, "ohm", ""), (47000, "ohm", "")),
            ((1000000, "ohm", ""), (1500000, "ohm", ""), (860000, "ohm", ""), (8600000, "ohm", "")),
            ((1.2, "ohm", ""), (3.33, "ohm", ""), (0.12, "ohm", ""), (1.234, "ohm", ""), (0.33, "ohm", ""), (3.3e-8, "ohm", "")),
            ((0.000010, "F", ""), (1e-3, "F", ""), (10e-3, "H", "")),
            ((1.5e-6, "F", ""), (33e-12, "F", ""), (100e-9, "F", "")),
            ((1, "ohm", ""), (0.1, "ohm", ""), (10, "ohm", ""), (100, "ohm", ""), (12.5, "ohm", "")),
            ((11, "R", ""), (120, "R", ""), (50, "R", "")),
            ((32768.0, "Hz", ""), (1e6,"Hz", ""), (12134,"Hz", "")),
        ]
        check = [
            ("1k", "1k5", "2k2", "2k21", "4k7", "47k"),
            ("1M", "1M5", "860k", "8M6"),
            ("1.2R", "3.33R", "0.12R", "1.234R", "0.33R", "33nohm"),
            ("10uF", "1mF", "10mH"),
            ("1.5uF", "33pF", "100nF"),
            ("1R", "0.1R", "10R", "100R", "12.5R"),
            ("11R", "120R", "50R"),
            ("32.768kHz", "1MHz", "12.134kHz"),
        ]

        for k, l in enumerate(test):
            for n, m in enumerate(l):
                l = lib.value_toStr(m, self.logger)
                self.assertTrue(l)

                print(l, "->", check[k][n])
                self.assertEqual(l, check[k][n])
                print("-" * 80)

    def test_mergeFileCommandLine(self):
        outfilename = os.path.join(self.temp_dir, 'cli_merged.xlsx')
        cmd  = ["python", "mergebom.py",
                '-o', 'cli_merged.xlsx',
                '-p', self.temp_dir,
                os.path.join( "test","Assembly","progettotest1","progettotest1.xlsx")]
        print(subprocess.check_output(cmd, stderr=subprocess.STDOUT))
        self.assertTrue(os.path.exists(outfilename), " ".join(cmd))
        os.remove(outfilename)

    def test_outFile(self):
        file_list = [
            "test/bom_uno.xls",
            "test/bom_due.xls",
            "test/bom_tre.xls",
            "test/bom_quattro.xls",
        ]

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()
        file_list = map(os.path.basename, file_list)
        ft = os.path.join(self.temp_dir, 'uno.xlsx')
        report.write_xls(
            d,
            file_list,
            self.config,
            ft,
            hw_ver="13",
            pcb="C",
            name="TEST")

    def test_parametri(self):
        import xlrd
        file_list = [
            os.path.join("test","Assembly","progettotest1","progettotest1.xlsx")
        ]

        r=report.Report(log_on_file=True, terminal=True, report_date=datetime.strptime('11/03/2018', '%d/%m/%Y'))
        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()
        file_list = map(os.path.basename, file_list)
        ft = os.path.join(self.temp_dir, 'due.xlsx')
        report.write_xls(
            d,
            file_list,
            self.config,
            ft,
            hw_ver="13",
            pcb="C",
            name="TEST")

        out = subprocess.check_output(["python", "mergebom.py",
                                      "-p", self.temp_dir,
                                      "-o", "due.xlsx",
                                      "-t", '11/03/2018',
                                      "--prj-hw-ver", "13",
                                      "--prj-name", "TEST",
                                      "--prj-pcb", "C"
                                      ,os.path.join("test","Assembly","progettotest1","progettotest1.xlsx")],
                                      stderr=subprocess.STDOUT)
        ft1=os.path.join(self.temp_dir, "due.xlsx")

        wb = xlrd.open_workbook(ft)
        data = []
        for s in wb.sheets():
            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    try:
                        curr = s.cell(row, col)
                    except IndexError:
                        continue

                    value = ""
                    try:
                        value = str(int(curr.value))
                    except (TypeError, ValueError):
                        value = unicode(curr.value)

                    values.append(value)
                data.append(values)
        uno=data

        wb = xlrd.open_workbook(ft1)
        data = []
        for s in wb.sheets():
            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    try:
                        curr = s.cell(row, col)
                    except IndexError:
                        continue

                    value = ""
                    try:
                        value = str(int(curr.value))
                    except (TypeError, ValueError):
                        value = unicode(curr.value)

                    values.append(value)
                data.append(values)

        risultato=True
        if len(uno) == len(data):
            for i,n in enumerate(uno):
                if not( uno[i] == data[i]):
                    risultato=False
        else:
            risultato=False
        self.assertTrue(risultato)

    def test_mergedFile(self):
        file_list = [
            os.path.join("test","bom-merged.xls"),
        ]

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.extra_data()
        print(d)
        self.assertEqual(len(d), 1)
        self.assertEqual(d[0]['project'], "test")
        self.assertEqual(d[0]['pcb_version'], "c")

    def test_stats(self):
        file_list = [
            os.path.join("test","bom-merged.xls"),
        ]

        self.logger.info(cfg.LOGO)
        m = MergeBom(file_list, self.config, logger=self.logger)
        m.merge()
        stats = m.statistics()
        self.logger.warning("File num: %s" % stats['file_num'])
        categories = self.config.categories()
        for i in stats.keys():
            if i in categories:
                self.logger.info(self.config.get(i, 'desc'))
                self.logger.info("%5.5s %5.5s" % (i, stats[i]))

        self.logger.warning("Total: %s" % stats['total'])

    def test_notPopulate(self):
        """
        Cerca i componenti da non montare
        """
        file_list = [
            os.path.join("test","bom-np.xls"),
        ]

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()
        stats = m.statistics()

        file_list = map(os.path.basename, file_list)
        ft = os.path.join(self.temp_dir, 'uno.xlsx')
        report.write_xls(d, file_list, self.config, ft, hw_ver="13",
                             pcb="C", name="TEST", statistics=stats)

    def test_cliMerge(self):
        outfilename = os.path.join(self.temp_dir, "climerge_merged-R0.xlsx")
        cmd = ["python", "mergebom.py",
                   "-prx", "climerge",
                   "-hw", "0",
                   "-pv", "S",
                   "-n", "Test project",
                   "-p", self.temp_dir,
                   "-e",
                   os.path.join("test","cli-merge0.xlsx"),
                   os.path.join("test","cli-merge1.xlsx")]
        print()
        print("%s" % " ".join(cmd))
        print(subprocess.check_output(cmd, stderr=subprocess.STDOUT))

        print(outfilename)
        self.assertTrue(os.path.isfile(outfilename))
        os.remove(outfilename)

    def test_cliMergeDiff(self):
        outfilename = os.path.join(self.temp_dir, "cli-diff-R23.xlsx")
        cmd = ["python", "mergebom.py", "-o", "cli-diff-R23.xlsx",
                       "-hw", "23",
                       "-pv", "T",
                       "-p", self.temp_dir,
                       "-d",
                       os.path.join("test","cli-merge-diff0.xlsx"),
                       os.path.join("test","cli-merge-diff1.xlsx")]
        print()
        print(" ".join(cmd))
        print(subprocess.check_output(cmd, stderr=subprocess.STDOUT))

        print(outfilename)
        self.assertTrue(os.path.isfile(outfilename), "Merged diff File not generated" )
        os.remove(outfilename)

        outfilename = os.path.join(self.temp_dir, "cli-diff_merged.xlsx")
        cmd = ["python", "mergebom.py",
                           "-o", "cli-diff_merged.xlsx",
                           "-d",
                           "-p", self.temp_dir,
                           os.path.join("test","diff_test_old.xlsx"),
                           os.path.join("test","diff_test_new.xlsx")]
        print()
        print(" ".join(cmd))
        print(subprocess.check_output(cmd, stderr=subprocess.STDOUT))

        self.assertTrue(os.path.isfile(outfilename), "Merged diff File not generated")

    def test_cliMergeGlob(self):
        outfilename = os.path.join(self.temp_dir, "cli-mergedGlob-R53.xlsx")
        cmd = ["python", "mergebom.py", "-hw", "53", "-pv", "O",
                       "-n", "Test project glob",
                       "-o", "cli-mergedGlob-R53.xlsx",
                       "-p", self.temp_dir,
                       os.path.join("test", "diff_test_old.xlsx")]
        print()
        print(" ".join(cmd))
        print(subprocess.check_output(cmd, stderr=subprocess.STDOUT))

        self.assertTrue(os.path.isfile(outfilename), "Merged File not generated")
        os.remove(outfilename)

    def test_categoryGroup(self):
        file_list = [
            "test/category_sysexit.xls",
        ]

        with self.assertRaises(SystemExit):
            print("Test should make sys Exit for not founded key.")
            m = MergeBom(file_list, self.config, logger=self.logger)
            d = m.merge()

        file_list = [
            "test/category.xls",
        ]

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()
        key = d.get('F', None)
        self.assertIsNotNone(key, "Check Category non found..")

    def test_otherColumn(self):
        file_list = [
            "test/column.xlsx",
        ]

        check = {'C': [
                    [1, 1, u'C1', '10nF', '805', u'x5r', u'123', u'cose varie note', '',''],
                    [1, 1, u'C2', '100nF', '805', u'x5r', u'789', '', 'cde', 'cde'],
                    [10, 10, u'C3, C4, C5, C6, C7, C8, C9, C10, C11, C12', '100nF',
                     '805', u'x7r', u'123; 456', u'altro', u'abc', u'code abc'],
                ],
                'U': [
                    [2, 2, u'u2, u3', u'lm2902', u'soic', u'Op-amp', u'aa',
                     u'Aa-bb; bb', u'cc', u'dd'],
                    [1, 1, u'u1', u'lm75', u'soic', u'temp', u'uno', u'due',
                     u'tre', u'quattro']
                 ]
        }
        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()

        print()
        dump(d)
        for i in d.keys():
            for n, j in enumerate(d[i]):
                print("T >", j)
                print("C <", check[i][n])
                self.assertEqual(len(j), len(check[i][n]))
                for m, c in enumerate(check[i][n]):
                    print("T >", c)
                    print("C <", j[m])
                    self.assertEqual(c, j[m])

        outfilename = os.path.join(self.temp_dir, "extra_column_merged-R0.xlsx")
        cmd = ["python",
               "mergebom.py",
               "-prx", "extra_column",
               "-p", self.temp_dir,
               "-hw", "0",
               "-pv", "S",
               "-n", "Test project",
               "test/column.xlsx"]
        print(" ".join(cmd))
        print(subprocess.check_output(cmd, stderr=subprocess.STDOUT))

        self.assertTrue(os.path.isfile(outfilename), "Merged File not generated")

    def test_cliCSV(self):
        inputfilename = os.path.join("test", "Assembly", "progettotest1", "progettotest1.csv")
        outfilename = os.path.join(".", "bom-merged.xlsx")
        cmd =["python", "mergebom.py", "--csv",
              "-hw", "77",
              "-pv", "X",
              "-n", "Test project CVS",
              inputfilename]
        print(" ".join(cmd))
        print(subprocess.check_output(cmd, stderr=subprocess.STDOUT))

        self.assertTrue(os.path.isfile(outfilename), "No mergefile generated")
        os.remove(outfilename)

    def test_extracPrj(self):
        wk_file = os.path.join("test", "wktest", "camera.DsnWrk")
        a = cfg.extrac_projects(wk_file)
        ck = [
            ('5mp-sensor', os.path.join('5mp-sensor', '5mp-sensor.PrjPCB')),
            ('camera-core', os.path.join('camera-core', 'camera-core.PrjPCB')),
            ('18mp-sensor', os.path.join('18mp-sensor', '18mp-sensor.PrjPCB')),
            ('imx8m_evk', os.path.join('imx8m_evk', 'imx8_evk.PrjPcb')),
            ('camera', 'camera.PrjPcb')
        ]
        print(a)
        self.assertEqual(a, ck)

    def test_prjParam(self):
        root = os.path.join("test", "wktest")
        tst = [
            ('5mp-sensor', os.path.join(root, '5mp-sensor', '5mp-sensor.PrjPCB')),
            ('camera-core', os.path.join(root, 'camera-core', 'camera-core.PrjPCB')),
            ('18mp-sensor', os.path.join(root, '18mp-sensor', '18mp-sensor.PrjPCB')),
            ('imx8m_evk', os.path.join(root, 'imx8m_evk', 'imx8_evk.PrjPcb')),
            ('camera', 'camera.PrjPcb')
        ]

        ch = [
            ['5mp-sensor',
                {
                    'prj_date': '01/06/2018',
                    'prj_hw_ver': '0',
                    'prj_license': 'Copyright company spa',
                    'prj_name': 'Sensor Shield 5Mp',
                    'prj_name_long': 'Camera sensor shield 5Mp AR0521',
                    'prj_pcb': 'A', 'prj_prefix': '',
                    'prj_status': 'Prototype'
                    }
             ],
            ['camera-core', {'prj_date': '01/06/2018', 'prj_hw_ver': '0', 'prj_license': 'Copyright company spa', 'prj_name': 'Camera Core', 'prj_name_long': 'Camera core imx8', 'prj_pcb': 'A', 'prj_pn': '-', 'prj_prefix': '', 'prj_status': 'Prototype'}],
            ['18mp-sensor', {'prj_status': 'Prototype', 'prj_prefix': '', 'prj_pcb': 'A', 'prj_name_long': 'Camera sensor shield 5Mp AR1820', 'prj_name': 'Sensor Shield 18Mp', 'prj_license': 'Copyright company spa', 'prj_hw_ver': '0', 'prj_date': '01/06/2018'}],
            ['imx8m_evk', {}],
            [],
        ]
        for n, i in enumerate(tst):
            a = cfg.get_parameterFromPrj(i[0], i[1])
            print(a)
            self.assertEqual(ch[n], a)

    def test_fileList(self):
        root = os.path.join("test", "wktest", "Assembly")
        tst = [
            ('5mp-sensor', os.path.join(root, '5mp-sensor', '5mp-sensor.PrjPCB')),
            ('camera-core', os.path.join(root, 'camera-core', 'camera-core.PrjPCB')),
            ('18mp-sensor', os.path.join(root, '18mp-sensor', '18mp-sensor.PrjPCB')),
            ('imx8m_evk', os.path.join(root, 'imx8m_evk', 'imx8_evk.PrjPcb')),
            ('camera', 'camera.PrjPcb')
        ]
        ck = [
            ['5mp-sensor', []],
            ['camera-core', ['test/wktest/Assembly/camera-core/bom-camera-core.xlsx']],
            ['18mp-sensor', []],
            ['imx8m_evk', []],
            ['camera', []]
        ]
        for n, i in enumerate(tst):
            a = cfg.find_bomfiles(root, i[0], False)
            print(a)
            self.assertEqual(ck[n], a)

if __name__ == "__main__":

    from optparse import OptionParser

    parser = OptionParser()
    parser.add_option(
        "-v",
        "--verbose",
        dest="verbose",
        default='2',
        help="Output verbosity")
    (options, args) = parser.parse_args()

    suite = unittest.TestSuite()
    suite.addTest(TestMergeBom("test_extracPrj"))
    suite.addTest(TestMergeBom("test_prjParam"))
    suite.addTest(TestMergeBom("test_fileList"))
    suite.addTest(TestMergeBom("test_altiumWorkspace"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_group"))
    suite.addTest(TestMergeBom("test_led"))
    suite.addTest(TestMergeBom("test_rele"))
    suite.addTest(TestMergeBom("test_groupFmt"))
    suite.addTest(TestMergeBom("test_diff"))
    suite.addTest(TestMergeBom("test_orderRef"))
    suite.addTest(TestMergeBom("test_valueToFloat"))
    suite.addTest(TestMergeBom("test_floatToValue"))
    suite.addTest(TestMergeBom("test_outFile"))
    suite.addTest(TestMergeBom("test_parametri"))
    suite.addTest(TestMergeBom("test_mergedFile"))
    suite.addTest(TestMergeBom("test_stats"))
    suite.addTest(TestMergeBom("test_notPopulate"))
    suite.addTest(TestMergeBom("test_otherColumn"))
    suite.addTest(TestMergeBom("test_categoryGroup"))
    suite.addTest(TestMergeBom("test_cliMerge"))
    suite.addTest(TestMergeBom("test_cliMergeDiff"))
    suite.addTest(TestMergeBom("test_cliMergeGlob"))
    suite.addTest(TestMergeBom("test_cliCSV"))
    suite.addTest(TestMergeBom("test_mergeFileCommandLine"))
    unittest.TextTestRunner(
        stream=sys.stdout,
        verbosity=2).run(suite)





