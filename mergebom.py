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
import re
import datetime
import tempfile
import shutil
from lib import report, cfg, lib

# Exchange data layout after file import
FILENAME = 0
QUANTITY = 1
DESIGNATOR = 2
DESCRIPTION = 3
COMMENT = 4
FOOTPRINT = 5
FARNELL = 6


class MergeBom (object):
    def __init__(self, list_bom_files, config, logger=None, terminal=True):
        """
        Data structure

        [
        file_name: [
        designator : [ all ]
        designator : [ all ]
        ...
        designator : [ all ]
        ],
        ...
        file_name: [
        designator : [ all ]
        designator : [ all ]
        ...
        designator : [ all ]
        ],
        ]

        """
        self.files = {}
        self.table_list = []
        self.extra_keys = []
        self.stats = {}
        self.terminal = terminal

        if logger is None:
            print
            print "Error you should specify a logger class"
            print
            sys.exit(1)

        self.logger = logger

        self.config = config
        self.categories = self.config.categories()

        self.stats['file_num'] = 0
        for index_file, file_name in enumerate(list_bom_files):
            self.stats['file_num'] += 1
            self.logger.warning(file_name)
            wb, data = report.read_xls(file_name)
            n = os.path.basename(file_name)

            if n in self.files:
                n = "%s-%s" % (index_file, n)
            self.files[n] = index_file

            # Get all header keys
            header = {}
            extra_keys = {}
            for row in data:
                for n, item in enumerate(row):

                    # search header to import corretly all data
                    if item.lower() in cfg.VALID_KEYS:
                        header[item.lower()] = n

                    # search extra data, like project name and revision
                    k = ""
                    try:
                        k, v = item.lower().split(":")
                    except ValueError:
                        continue
                    if k in cfg.EXTRA_KEYS:
                        extra_keys[k] = v.replace(' ', '')

            self.extra_keys.append(extra_keys)

            try:
                designator = header['designator']
                comment = header['comment']
                footprint = header['footprint']
                description = header['description']
                farnell = header['farnell']
            except KeyError as e:
                self.logger.error("No key header found! [%s]\n" % e)
                self.logger.warning("Valid are:")
                for i in cfg.VALID_KEYS:
                    self.logger.warning(" %s" % i)

                sys.exit(1)

            table_dict = {}
            for row in data:
                # skip header
                if row[designator].lower() == 'designator':
                    continue

                if filter(lambda x: x, row):
                    # Fix designator field, we want all designator separated
                    # by comma followed by space.
                    # In this way excel could resize row cell correctly
                    if re.findall("\S,[\S]+", row[designator]):
                        row[designator] = row[designator].replace(",", ", ")

                        # Explode designator field to have one component for
                        # line
                    d = row[designator].split(',')
                    for reference in d:
                        r = reference.replace(' ', '')
                        if r:
                            table_dict[r] = [
                                os.path.basename(file_name),
                                1,  # one item for a row
                                r,  # designator
                                row[description],
                                row[comment],
                                row[footprint],
                                row[farnell]
                            ]

            self.table_list.append(table_dict)

    def table_data(self):
        return self.table_list

    def extra_data(self):
        return self.extra_keys

    def group(self):
        self.grouped_items = {}
        for table_dict in self.table_list:
            for designator in table_dict.keys():
                # Grop found designator componets by its category
                c = re.search('^[a-zA-Z_]{1,3}', designator)
                group_key = ''
                if c is not None:
                    group_key = self.config.check_category(c.group().upper())

                    if not group_key:
                        self.logger.warning("WARNING!! KEY SKIPPED [%s]\n" % group_key)
                        continue

                    if group_key is None:
                        self.logger.error("GROUP key not FOUND!\n")
                        self.logger.error( "%s, %s, %s\n" % (c.group(), designator,
                             table_dict[designator]))
                        sys.exit(1)

                    if group_key in self.grouped_items:
                        self.grouped_items[group_key].append(
                            table_dict[designator])
                    else:
                        self.grouped_items[group_key] = [
                            table_dict[designator]]
                else:
                    self.logger.error("GROUP key not FOUND!\n")
                    self.logger.error(designator)
                    sys.exit(1)

    def table_grouped(self):
        return self.grouped_items

    def count(self):
        """
        grouped items format
        'U': ['bom_due.xls', 1, u'U3000', u'Lan transformer', u'LAN TRANFORMER WE 7490100111A', u'LAN_TR_1.27MM_SMD']
        'U': ['test.xlsx', 1, u'U2015', u'DC/DC switch converter', u'LM22671MR-ADJ', u'SOIC8_PAD']
        'U': ['test.xlsx', 1, u'U2002', u'Dual Low-Power Operational Amplifier', u'LM2902', u'SOIC14']
        'U': ['bom_uno.xls', 1, u'U7', u'Temperature sensor', u'LM75BIM', u'SOIC8']
        """

        self.table = {}

        # Finally Table layout
        self.TABLE_TOTALQTY = 0
        self.TABLE_DESIGNATOR = len(self.files) + 1
        self.TABLE_COMMENT = self.TABLE_DESIGNATOR + 1
        self.TABLE_FOOTPRINT = self.TABLE_COMMENT + 1
        self.TABLE_DESCRIPTION = self.TABLE_FOOTPRINT + 1
        self.TABLE_FARNELL = self.TABLE_DESCRIPTION + 1

        self.stats['total'] = 0
        for category in self.categories:
            if category in self.grouped_items:
                tmp = {}
                self.stats[category] = 0
                for item in self.grouped_items[category]:

                    # Fix Designator
                    if category in ["R", "C", "L", "Y"]:
                        tmp_comment = lib.value_toFloat(item[COMMENT], category, self.logger)
                        item[COMMENT] = lib.value_toStr(tmp_comment, self.logger)

                    # Fix Not poluate string in list
                    for rexp in cfg.NOT_POPULATE_KEY:
                        item[COMMENT] = re.sub(rexp, 'NP ', item[COMMENT])

                    if category == 'J':
                        # Avoid merging for NP componets
                        skip_merge = False
                        m = re.findall(cfg.NP_REGEXP, item[COMMENT])
                        if m:
                            skip_merge = True
                            self.logger.error(
                                "Not Populate connector, leave unmerged..[%s] [%s] match%s\n" %
                                (item[COMMENT], item[DESIGNATOR], m))
                            item[COMMENT] = "NP Connector"

                        if skip_merge:
                            key = item[DESCRIPTION] + \
                                item[COMMENT] + item[FOOTPRINT]
                        else:
                            key = item[DESCRIPTION] + item[FOOTPRINT]
                            item[COMMENT] = "Connector"

                        self.logger.warning("Merged key: %s (%s)\n" %
                            (key, item[COMMENT]))

                    if category == 'D' and "LED" in item[FOOTPRINT]:
                        key = item[DESCRIPTION] + item[FOOTPRINT]
                        item[COMMENT] = "LED"
                        self.logger.warning("Merged key: %s (%s)\n" % (key, item[COMMENT]))
                    
                    

                    if category == 'S' and "TACTILE" in item[FOOTPRINT]:
                        key = item[DESCRIPTION] + item[FOOTPRINT]
                        item[COMMENT] = "Tactile Switch"
                        self.logger.warning("Merged key: %s (%s)\n" % (key, item[COMMENT]))

                    elif category == 'U' and re.findall("rele|relay", item[DESCRIPTION].lower()):
                        key = item[DESCRIPTION] + item[FOOTPRINT]
                        item[COMMENT] = u"Relay, Rele'"
                        self.logger.warning("Merged key: %s (%s)\n" % (key, item[COMMENT]))
                    else:
                        key = item[DESCRIPTION] + \
                            item[COMMENT] + item[FOOTPRINT]

                    # print key
                    # print "<<", item[DESIGNATOR]
                    self.stats[category] += 1
                    self.stats['total'] += 1

                    # First colum is total Qty
                    curr_file_index = self.files[
                        item[FILENAME]] + self.TABLE_TOTALQTY + 1

                    if key in tmp:
                        # print "UPD", tmp[key], curr_file_index,
                        # item[FILENAME]
                        tmp[key][self.TABLE_TOTALQTY] += item[QUANTITY]
                        tmp[key][curr_file_index] += item[QUANTITY]
                        tmp[key][
                            self.TABLE_DESIGNATOR] += ", " + item[DESIGNATOR]
                        tmp[key][self.TABLE_DESIGNATOR] = lib.order_designator(
                            tmp[key][self.TABLE_DESIGNATOR], self.logger)
                    else:
                        row = [item[QUANTITY]] + \
                            [0] * len(self.files) + \
                            [
                                item[DESIGNATOR],
                                item[COMMENT],
                                item[FOOTPRINT],
                                item[DESCRIPTION],
                                item[FARNELL],
                            ]

                        row[curr_file_index] = item[QUANTITY]
                        tmp[key] = row
                        # print "NEW", tmp[key], curr_file_index,
                        # item[FILENAME]

                self.table[category] = tmp.values()

    def statistics(self):
        return self.stats

    def merge(self):
        self.group()
        self.count()
        for category in self.categories:
            if category in self.table:
                for n, item in enumerate(self.table[category]):
                    self.table[category][n][self.TABLE_DESIGNATOR] = \
                        lib.order_designator(item[self.TABLE_DESIGNATOR], self.logger)

                # Convert all designator in a number to be ordered
                if category in ["R", "C", "L", "Y"]:
                    for m in self.table[category]:
                        m[self.TABLE_COMMENT] = lib.value_toFloat(
                            m[self.TABLE_COMMENT], category, self.logger)
                        # print m[COMMENT], key

                self.table[category] = sorted(
                    self.table[category], key=lambda x: x[
                        self.TABLE_COMMENT])

                # Convert all ORDERED designator in a numeric format
                if category in ["R", "C", "L", "Y"]:
                    for m in self.table[category]:
                        m[self.TABLE_COMMENT] = lib.value_toStr(
                            m[self.TABLE_COMMENT], self.logger)
                        # print m[self.TABLE_COMMENT], category

        return self.table

    def diff(self):
        if len(self.table_list) > 2:
            self.logger.error("To much file ti compare!\n")
            sys.exit(1)
        diff = {}
        self.logger.warning("%s\n" % self.files.items())
        for i in self.files.items():
            if i[1] == 0:
                fA = i[0]
                A = self.table_list[0]
            if i[1] == 1:
                fB = i[0]
                B = self.table_list[1]

        self.logger.warning("A:%s B:%s\n" % (fA, fB))

        for k in A.keys():
            if k in B:

                c = re.search('^[a-zA-Z_]{1,3}', A[k][DESIGNATOR])
                category = ''
                if c is not None:
                    category = c.group().upper()

                if category == 'J':
                    la = [A[k][DESIGNATOR], A[k][FOOTPRINT]]
                    lb = [B[k][DESIGNATOR], B[k][FOOTPRINT]]

                    self.logger.warning("Merged key: %s (%s)\n" % (k, A[k][COMMENT]))

                if category == 'D' and "LED" in A[k][FOOTPRINT]:
                    la = [A[k][DESIGNATOR], A[k][FOOTPRINT]]
                    lb = [B[k][DESIGNATOR], B[k][FOOTPRINT]]

                    self.logger.warning("Merged key: %s (%s)\n" % (k, A[k][COMMENT]))
                else:
                    la = A[k][1:]
                    lb = B[k][1:]

                if la != lb:
                    diff[k] = (A[k], B[k])

                del B[k]
            else:
                diff[k] = (A[k], [fB] + ['-'] * (len(A[k]) - 1))

            del A[k]

        for k in B.keys():
            diff[k] = ([fA] + ['-'] * (len(B[k]) - 1), B[k])

        return diff

if __name__ == "__main__":
    import glob
    from optparse import OptionParser

    file_list = []
    parser = OptionParser()
    parser.add_option("-c", "--merge-cfg", dest="merge_cfg",
                      default=None, help="MergeBOM configuration file.")
    parser.add_option("-p", "--search-dir", dest="search_dir",
                      default='./', help="BOM to merge search path.")
    parser.add_option("-o", "--out-filename", dest="out_filename",
                      default='merged_bom.xlsx', help="Out file name")
    parser.add_option("-v", "--version-file", dest="version_file",
                      default='version.txt', help="Version file.")
    parser.add_option(
        "-l",
        "--log-on-file",
        dest="log_on_file",
        default=True,
        action="store_true",
        help="List all project name from version file.")
    parser.add_option(
        "-s",
        "--sub_search_dir",
        dest="sub_search_dir",
        default="Assembly",
        help="Default name for BOM store directory in version file mode.")
    parser.add_option(
        "-d",
        "--diff",
        dest="diff",
        action="store_true",
        default=False,
        help="Generate diff from two specified BOMs")
    parser.add_option(
        "-m",
        "--replace-merged",
        dest="replace_merged",
        action="store_false",
        default=True,
        help="Remove all boms files after merge (only when using version file)")
    parser.add_option("-r", "--bom-revision", dest="bom_rev",
                      default=None, help="Hardware BOM revision")
    parser.add_option("-w", "--bom-pcb-revision", dest="bom_pcb_ver",
                      default=None, help="PCB Revision")
    parser.add_option("-n", "--bom-prj-name", dest="bom_prj_name",
                      default=None, help="Project names.")
    parser.add_option(
        "-t",
        "--bom-date",
        dest="bom_prj_date",
        default=datetime.datetime.today().strftime("%d/%m/%Y"),
        help="Project date.")

    (options, args) = parser.parse_args()

    # Get logger
    logger = report.Report(log_on_file=options.log_on_file, terminal=True)
    logger.write_logo()

    # Load default Configuration for merging
    config = cfg.CfgMergeBom(options.merge_cfg)

    # The user specify file to merge
    file_list = args
    logger.info("Merge BOM file..\n")
    if file_list:
        logger.info("Merge Files:\n")
        logger.info("%s\n" % args)

    if not file_list and os.path.isfile(options.version_file):
        logger.warning("BOM file not found..\n")

        # search version file
        logger.info("Search version file [%s] in [%s]:\n" %
            (options.version_file,
             options.search_dir))

        # Get bom file list from version.txt
        """
        Search BOM files in Project directory.
        By defaul the script search in "sub_search_dir" (Assembly) directory all boms like:
        - the folder name contained in "sub_search_dir"/section_name_in_version_file
        - the bom file named as section_name_in_version_file in "sub_search_dir"
        """
        bom_info = cfg.cfg_version(options.version_file)
        for prj in bom_info.keys():
            file_list = []
            glob_path = os.path.join(
                options.search_dir, options.sub_search_dir, prj)
            search_glob = ["*.xls", "*.xlsx"]
            if not os.path.exists(glob_path):
                glob_path = os.path.join(
                    options.search_dir, options.sub_search_dir)
                search_glob = ["*" + prj + "*.xls", "*" + prj + "*.xlsx"]

            for i in search_glob:
                file_list += glob.glob(os.path.join(glob_path, i))

            if not file_list:
                logger.error("No BOM file to merge\n")
                sys.exit(1)

            m = MergeBom(file_list, config, logger=logger)
            d = m.merge()
            stats = m.statistics()

            tmpdir = tempfile.gettempdir()
            outfilename = os.path.join(
                tmpdir, os.path.basename(
                    options.out_filename))
            dstfilename = os.path.join(
                os.path.dirname(
                    options.out_filename),
                prj +
                "_" +
                os.path.basename(
                    options.out_filename))

            if options.replace_merged:
                outfilename = os.path.join(tmpdir, "bom-%s.xlsx" % prj)
                dstfilename = os.path.join(glob_path, "bom-%s.xlsx" % prj)

            name = bom_info[prj]["name"]
            hw_ver = bom_info[prj]["hw_ver"]
            pcb_ver = bom_info[prj]["pcb_ver"]

            report.write_xls(d,
                             map(os.path.basename, file_list),
                             config,
                             outfilename,
                             hw_ver=hw_ver,
                             pcb_ver=pcb_ver,
                             project=name,
                             statistics=stats)

            if os.path.isfile(outfilename):
                if options.replace_merged:
                    for i in file_list:
                        os.remove(i)

                shutil.copy(outfilename, dstfilename)
                logger.info("Generated %s Merged BOM file.\n" % dstfilename)

        sys.exit(0)

    if not file_list:
        logger.warning(
            "No BOM specified to merge..\n")
        # First merge all xlsx file in search directory
        logger.info("Find in search_dir[%s] all bom files:\n" % options.search_dir)
        file_list = glob.glob(os.path.join(options.search_dir, '*.xls'))
        file_list += glob.glob(os.path.join(options.search_dir, '*.xlsx'))

    # File list empty, so there aren't BOM files to merge, raise error to user.
    if not file_list:
        logger.warning("No version file found..\n")
        parser.print_help()
        sys.exit(1)

    # The user set diff mode, but this works only for two bom.
    if options.diff and len(file_list) != 2:
        logger.error("In diff mode you should specify only 2 BOMs [%s].\n" % len(file_list))
        parser.print_help()
        sys.exit(1)

    m = MergeBom(file_list, config, logger=logger)
    file_list = map(os.path.basename, file_list)

    if options.diff:
        d = m.diff()
        l = m.extra_data()
        report.write_xls(
            d,
            file_list,
            config,
            options.out_filename,
            diff=True,
            extra_data=l)
        sys.exit(0)

    d = m.merge()
    stats = m.statistics()

    if options.bom_rev is None \
            or options.bom_pcb_ver is None \
            or options.bom_prj_name is None:
        logger.error("\nYou should specify some missing parameter:\n")
        logger.error("- project name: %s\n" % options.bom_prj_name)
        logger.error("- hw verion: %s\n" % options.bom_rev)
        logger.error("- pcb version: %s\n" % options.bom_pcb_ver)
        sys.exit(1)


    report.write_xls(
        d,
        file_list,
        config,
        options.out_filename,
        hw_ver=options.bom_rev,
        pcb_ver=options.bom_pcb_ver,
        project=options.bom_prj_name,
        statistics=stats)
