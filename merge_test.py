

        for i in d.keys():
            for n, j in enumerate(d[i]):
                print "T >", j
                print "C <", check[n]
                for m, c in enumerate(check[n]):
                    print "T >", c
                    print "C <", j[m]
                    self.assertEqual(c, j[m])

    def test_groupFmt(self):
        file_list = [
            "test/bom-fmt.xls",
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
            "220h ohm @ 100mhz": (1, 14),
            "100nf": (1, 30),
            "6.8r": (1, 3),
            "ft232rq": (1, 1),
            "line filter (we 744232090)": (1, 2),
            "np  (220 ohm @ 100mhz)": (1, 1),
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
            #print "\"%s\": (%s, %s)," % (k, test_dict[k][0], test_dict[k][1])
            print "T > %20.20s | row count %5.5s | count %5.5s |" % (k, test_dict[k][0], test_dict[k][1])
            print "C < %20.20s | row count %5.5s | count %5.5s |" % (k, check[k][0], check[k][1])

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
            ("R", ("1.2R", "3.33R", "0.12R", "1.234R", "0.33R", "33nohm")),
            ("Y", ("32.768kHz", "1MHz", "12.134kHz")),
        ]
        checkv1 = [
            ((1e-10, "F"),),
            ((0.3, "ohm"), (1.1, "ohm"), (1.2, "ohm"), (1.8, "ohm")),
            ((0.3, "ohm"), (1.1, "ohm"), (1.2, "ohm"), (1.8, "ohm")),
            ((0.3, "ohm"), (1.0, "ohm"), (1.2, "ohm"), (10.0, "ohm"), (1000.0, "ohm"), (1500.0, "ohm"), (2200.0, "ohm")),
            ((10e-12, "F"), (100e-9, "F"), (0.1e-6, "F"), (1e-6, "F"), (2.2e-6, "F"), (47e-6, "F"), (1.0, "F")),
            ((10e-12, "H"), (1e-9, "H"), (2.2e-6, "H"), (47e-6, "H"), (1.0, "H")),
            ((3.3e-8, "ohm"), (0.12, "ohm"), (0.33, "ohm"), (1.2, "ohm"), (1.234, "ohm"), (3.33, "ohm")),
            ((12134.0, "Hz"), (32768.0, "Hz"), (1e6, "Hz")),
        ]

        print
        for k, m in enumerate(test):
            l = []
            for mm in m[1]:
                a, b, c = lib.value_toFloat(mm, m[0], self.logger)
                l.append((a, b))

            l = sorted(l, key=lambda x: x[0])
            for n, i in enumerate(l):
                print i, "->", checkv1[k][n]
                self.assertTrue(str(i[0]) == str(checkv1[k][n][0]))
                self.assertEqual(i[1], checkv1[k][n][1])
                print "-" * 80

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

                print l, "->", check[k][n]
                self.assertEqual(l, check[k][n])
                print "-" * 80

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
            pcb_ver="C",
            project="TEST")

    def test_mergedFile(self):
        file_list = [
            "test/bom-merged.xls",
        ]

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.extra_data()
        print d
        self.assertEqual(len(d), 1)
        self.assertEqual(d[0]['project'], "test")
        self.assertEqual(d[0]['pcb_version'], "c")

    def test_stats(self):
        file_list = [
            "test/bom-merged.xls",
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
            "test/bom-np.xls",
        ]

        m = MergeBom(file_list, self.config, logger=self.logger)
        d = m.merge()
        stats = m.statistics()

        file_list = map(os.path.basename, file_list)
        ft = os.path.join(self.temp_dir, 'uno.xlsx')
        report.write_xls(d, file_list, self.config, ft, hw_ver="13",
                             pcb_ver="C", project="TEST", statistics=stats)

    def test_cliMerge(self):
        outfilename = os.path.join(self.temp_dir, "cli-merged.xlsx")
        out = subprocess.check_output(["python",
                                       "mergebom.py",
                                       "-o",
                                       outfilename,
                                       "-r", "0",
                                       "-w", "S",
                                       "-n", "Test project",
                                       "test/cli-merge0.xlsx",
                                       "test/cli-merge1.xlsx"],
                                      stderr=subprocess.STDOUT)

        print out
        self.assertTrue(
            os.path.isfile(outfilename),
            "Merged File not generated")
        os.remove(outfilename)

    def test_cliMergeDiff(self):
        outfilename = os.path.join(self.temp_dir, "cli-diff-merged.xlsx")
        out = subprocess.check_output(["python",
                                       "mergebom.py",
                                       "-o",
                                       outfilename,
                                       "-r", "23",
                                       "-w", "T",
                                       "-n", "Test project diff",
                                       "-d",
                                       "test/cli-merge-diff0.xlsx",
                                       "test/cli-merge-diff1.xlsx"],
                                      stderr=subprocess.STDOUT)

        print out
        self.assertTrue(
            os.path.isfile(outfilename),
            "Merged diff File not generated")
        os.remove(outfilename)

        retcode = None
        try:
            out = subprocess.check_call(["python",
                                         "mergebom.py",
                                         "-o",
                                         outfilename,
                                         "-d",
                                         "test/cli-merge-diff0.xlsx",
                                         "test/cli-merge-diff1.xlsx",
                                         "test/cli-merge-diff2.xlsx"],
                                        stderr=subprocess.STDOUT)
        except subprocess.CalledProcessError as e:
            print e
            retcode = e.returncode

        self.assertEqual(retcode, 1)

        out = subprocess.check_call(["python",
                                     "mergebom.py",
                                     "-o",
                                     outfilename,
                                     "-d",
                                     "test/diff_test_old.xlsx",
                                     "test/diff_test_new.xlsx"],
                                    stderr=subprocess.STDOUT)
        self.assertTrue(
            os.path.isfile(outfilename),
            "Merged diff File not generated")

    def test_cliMergeGlob(self):
        outfilename = os.path.join(self.temp_dir, "cli-mergedGlob.xlsx")
        out = subprocess.check_output(["python", "mergebom.py",
                                       "-r", "53",
                                       "-w", "O",
                                       "-n", "Test project glob",
                                       "-o", outfilename, "-p", "test/glob/"],
                                      stderr=subprocess.STDOUT)

        print out
        self.assertTrue(
            os.path.isfile(outfilename),
            "Merged File not generated")
        os.remove(outfilename)

    def test_categoryGroup(self):
        file_list = [
            "test/category_sysexit.xls",
        ]

        with self.assertRaises(SystemExit):
            print "Test should make sys Exit for not founded key."
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

        # file_list = map(os.path.basename, file_list)
        # stats = report.write_xls(h, d, file_list, "/tmp/uno.xls")

        print
        dump(d)
        for i in d.keys():
            for n, j in enumerate(d[i]):
                print "T >", j
                print "C <", check[i][n]
                self.assertEqual(len(j), len(check[i][n]))
                for m, c in enumerate(check[i][n]):
                    print "T >", c
                    print "C <", j[m]
                    self.assertEqual(c, j[m])

        outfilename = os.path.join(self.temp_dir, "extra_column.xlsx")
        out = subprocess.check_output(["python",
                                       "mergebom.py",
                                       "-o",
                                       outfilename,
                                       "-r", "0",
                                       "-w", "S",
                                       "-n", "Test project",
                                       "test/column.xlsx"],
                                       stderr=subprocess.STDOUT)

        print out
        self.assertTrue(
            os.path.isfile(outfilename),
            "Merged File not generated")
        os.remove(outfilename)

    def test_cliCSV(self):
        outfilename = os.path.join(self.temp_dir, "csv_test.xlsx")
        out = subprocess.check_output(["python", "mergebom.py",
                                       "--csv",
                                       "-r", "77",
                                       "-w", "X",
                                       "-n", "Test project CVS",
                                       "-o", outfilename, "test/test.csv"],
                                      stderr=subprocess.STDOUT)

        print out
        self.assertTrue(
            os.path.isfile(outfilename),
            "Merged File not generated")
        os.remove(outfilename)


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
    print args

    suite = unittest.TestSuite()
    suite.addTest(TestMergeBom("test_import"))
    suite.addTest(TestMergeBom("test_group"))
    suite.addTest(TestMergeBom("test_led"))
    suite.addTest(TestMergeBom("test_rele"))
    suite.addTest(TestMergeBom("test_diff"))
    suite.addTest(TestMergeBom("test_orderRef"))
    suite.addTest(TestMergeBom("test_outFile"))
    suite.addTest(TestMergeBom("test_mergedFile"))
    suite.addTest(TestMergeBom("test_stats"))
    suite.addTest(TestMergeBom("test_valueToFloat"))
    suite.addTest(TestMergeBom("test_floatToValue"))
    suite.addTest(TestMergeBom("test_notPopulate"))
    suite.addTest(TestMergeBom("test_cliMerge"))
    suite.addTest(TestMergeBom("test_cliMergeDiff"))
    suite.addTest(TestMergeBom("test_cliMergeGlob"))
    suite.addTest(TestMergeBom("test_groupFmt"))
    suite.addTest(TestMergeBom("test_categoryGroup"))
    suite.addTest(TestMergeBom("test_otherColumn"))
    suite.addTest(TestMergeBom("test_cliCSV"))
    unittest.TextTestRunner(
        stream=sys.stdout,
        verbosity=options.verbose).run(suite)
