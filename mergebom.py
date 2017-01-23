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

from lib import *

# Exchange data layout after file import
FILENAME    = 0
QUANTITY    = 1
DESIGNATOR  = 2
DESCRIPTION = 3
COMMENT     = 4
FOOTPRINT   = 5

class MergeBom (object):
    def __init__(self, list_bom_files, cfg, handler=sys.stdout, terminal=True):
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
        self.handler = handler
        self.terminal = terminal

        self.cfg = cfg
        self.categories = self.cfg.getCategories()

        self.stats['file_num'] = 0
        for index_file, file_name in enumerate(list_bom_files):
            self.stats['file_num'] += 1
            warning(file_name, self.handler, terminal=self.terminal)
            wb, data = read_xls(file_name)
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
                    if item.lower() in VALID_KEYS:
                        header[item.lower()] = n

                    # search extra data, like project name and revision
                    k = ""
                    try:
                        k, v = item.lower().split(":")
                    except ValueError:
                        continue
                    if k in EXTRA_KEYS:
                        extra_keys[k] = v.replace(' ', '')

            self.extra_keys.append(extra_keys)

            try:
                designator  = header['designator']
                comment     = header['comment']
                footprint   = header['footprint']
                description = header['description']
            except KeyError, e:
                error("No key header found! [%s]" % e, self.handler,
                      terminal=self.terminal)
                warning("Valid are:", self.handler, terminal=self.terminal)
                for i in VALID_KEYS:
                    warning("i", self.handler, terminal=self.terminal)

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

                   # Explode designator field to have one component for line
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
                                row[footprint]
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
                    group_key = self.cfg.checkGroup(c.group().upper())

                    if not group_key:
                        warning("WARNING!! KEY SKIPPED [%s]" % group_key,
                                self.handler, terminal=self.terminal)
                        continue

                    if group_key is None:
                        error("GROUP key not FOUND!", self.handler,
                              terminal=self.terminal)
                        error("%s, %s, %s" % (c.group(), designator, table_dict[designator]),
                              self.handler, terminal=self.terminal)
                        sys.exit(1)

                    if self.grouped_items.has_key(group_key):
                        self.grouped_items[group_key].append(table_dict[designator])
                    else:
                        self.grouped_items[group_key] = [table_dict[designator]]
                else:
                    error("GROUP key not FOUND!", self.handler, terminal=self.terminal)
                    error(designator, self.handler, terminal=self.terminal)
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
        self.TABLE_TOTALQTY    = 0
        self.TABLE_DESIGNATOR  = len(self.files) + 1
        self.TABLE_COMMENT     = self.TABLE_DESIGNATOR + 1
        self.TABLE_FOOTPRINT   = self.TABLE_COMMENT + 1
        self.TABLE_DESCRIPTION = self.TABLE_FOOTPRINT + 1

        self.stats['total'] = 0
        for category in self.categories:
            if self.grouped_items.has_key(category):
                tmp = {}
                self.stats[category] = 0
                for item in self.grouped_items[category]:

                    # Fix Not poluate string in list
                    for rexp in NOT_POPULATE_KEY:
                        item[COMMENT] = re.sub(rexp, 'NP ', item[COMMENT])

                    if category  == 'J':
                        # Avoid merging for NP componets
                        skip_merge = False
                        m = re.findall(NP_REGEXP, item[COMMENT])
                        if m:
                            skip_merge = True
                            error("Not Populate connector, leave unmerged..[%s] [%s] match%s" %
                                   (item[COMMENT], item[DESIGNATOR], m),
                                   self.handler, terminal=self.terminal)
                            item[COMMENT] = "NP Connector"

                        if skip_merge:
                            key = item[DESCRIPTION] + item[COMMENT] + item[FOOTPRINT]
                        else:
                            key = item[DESCRIPTION] + item[FOOTPRINT]
                            item[COMMENT] = "Connector"

                        warning("Merged key: %s (%s)" % (key, item[COMMENT]),
                                self.handler, terminal=self.terminal)

                    if category  == 'D' and "LED" in item[FOOTPRINT]:
                            key = item[DESCRIPTION] + item[FOOTPRINT]
                            item[COMMENT] = "LED"
                            warning("Merged key: %s (%s)" % (key, item[COMMENT]),
                                    self.handler, terminal=self.terminal)

                    if category  == 'S' and "TACTILE" in item[FOOTPRINT]:
                            key = item[DESCRIPTION] + item[FOOTPRINT]
                            item[COMMENT] = "Tactile Switch"
                            warning("Merged key: %s (%s)" % (key, item[COMMENT]),
                                    self.handler, terminal=self.terminal)

                    elif category  == 'U' and re.findall("rele|relay", item[DESCRIPTION].lower()):
                        key = item[DESCRIPTION] + item[FOOTPRINT]
                        item[COMMENT] = u"Relay, Rele'"
                        warning("Merged key: %s (%s)" % (key, item[COMMENT]),
                                self.handler, terminal=self.terminal)

                    else:
                        key = item[DESCRIPTION] + item[COMMENT] + item[FOOTPRINT]

                    #print key
                    #print "<<", item[DESIGNATOR]
                    self.stats[category] += 1
                    self.stats['total'] += 1

                    # First colum is total Qty
                    curr_file_index = self.files[item[FILENAME]] + self.TABLE_TOTALQTY + 1

                    if key in tmp:
                        #print "UPD", tmp[key], curr_file_index, item[FILENAME]
                        tmp[key][self.TABLE_TOTALQTY] += item[QUANTITY]
                        tmp[key][curr_file_index] += item[QUANTITY]
                        tmp[key][self.TABLE_DESIGNATOR] += ", " + item[DESIGNATOR]
                        tmp[key][self.TABLE_DESIGNATOR] = order_designator(tmp[key][self.TABLE_DESIGNATOR])
                    else:
                        row = [item[QUANTITY]] + \
                              [0] * len(self.files) + \
                              [
                                item[DESIGNATOR],
                                item[COMMENT],
                                item[FOOTPRINT],
                                item[DESCRIPTION]
                              ]

                        row[curr_file_index] = item[QUANTITY]
                        tmp[key] = row
                        #print "NEW", tmp[key], curr_file_index, item[FILENAME]

                self.table[category] = tmp.values()

    def statistics(self):
        return self.stats

    def merge(self):
        self.group()
        self.count()
        for category in self.categories:
            if self.table.has_key(category):
                for n, item in enumerate(self.table[category]):
                    self.table[category][n][self.TABLE_DESIGNATOR] = \
                            order_designator(item[self.TABLE_DESIGNATOR])

                if category in ["R", "C", "L", "Y"]:
                    for m in self.table[category]:
                        m[self.TABLE_COMMENT] = value_toFloat(m[self.TABLE_COMMENT], category)
                        #print m[COMMENT], key

                self.table[category] = sorted(self.table[category],
                                              key=lambda x: x[self.TABLE_COMMENT])

                if category in ["R", "C", "L", "Y"]:
                    for m in self.table[category]:
                        m[self.TABLE_COMMENT] = value_toStr(m[self.TABLE_COMMENT], category)
                        #print m[self.TABLE_COMMENT], category

        return self.table

    def diff(self):
        if len(self.table_list) > 2:
            error("To much file ti compare!", self.handler, terminal=self.terminal)
            sys.exit(1)
        diff = {}
        warning("%s" % self.files.items(), self.handler, terminal=self.terminal)
        for i in self.files.items():
            if i[1] == 0:
                fA = i[0]
                A = self.table_list[0]
            if i[1] == 1:
                fB = i[0]
                B = self.table_list[1]

        warning("A:%s B:%s" % (fA, fB), self.handler, terminal=self.terminal)

        for k in A.keys():
            if B.has_key(k):

                c = re.search('^[a-zA-Z_]{1,3}', A[k][DESIGNATOR])
                category = ''
                if c is not None:
                    category = c.group().upper()

                if category  == 'J':
                    la = [ A[k][DESIGNATOR], A[k][FOOTPRINT] ]
                    lb = [ B[k][DESIGNATOR], B[k][FOOTPRINT] ]

                    warning("Merged key: %s (%s)" % (k, A[k][COMMENT]), self.handler, terminal=self.terminal)

                if category  == 'D' and "LED" in A[k][FOOTPRINT]:
                    la = [ A[k][DESIGNATOR], A[k][FOOTPRINT] ]
                    lb = [ B[k][DESIGNATOR], B[k][FOOTPRINT] ]

                    warning("Merged key: %s (%s)" % (k, A[k][COMMENT]), self.handler, terminal=self.terminal)
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
    parser.add_option("-v", "--version-file", dest="version_file", default='version.txt', help="Version file.")
    parser.add_option("-o", "--out-filename", dest="out_filename", default='merged_bom.xlsx', help="Out file name")
    parser.add_option("-p", "--search-dir", dest="search_dir", default='./', help="BOM to merge search path.")
    parser.add_option("-d", "--diff", dest="diff", action="store_true", default=False, help="Generate diff from two specified BOMs")
    parser.add_option("-r", "--revision", dest="rev", default='0', help="Hardware BOM revision")
    parser.add_option("-w", "--pcb-revision", dest="pcb_ver", default='0', help="PCB Revision")
    parser.add_option("-n", "--prj-name", dest="prj_name", default='MyProject', help="Project names.")
    parser.add_option("-t", "--date", dest="prj_date", default=datetime.datetime.today().strftime("%d/%m/%Y"), help="Project date.")

    (options, args) = parser.parse_args()

    info(logo, sys.stdout, terminal=True, prefix="")



    # The user specify file to merge
    file_list = args
    info("Merge BOM file..", sys.stdout, terminal=True, prefix="")
    if file_list:
        info("Merge Files:", sys.stdout, terminal=True, prefix="")
        info("%s" % args, sys.stdout, terminal=True, prefix="")

    if not file_list:
        warning("No BOM specified to merge..", sys.stdout, terminal=True, prefix="")
        # First merge all xlsx file in search directory
        info("Find in search_dir[%s] all bom files:" % options.search_dir, \
                sys.stdout, terminal=True, prefix="")
        file_list = glob.glob(os.path.join(options.search_dir, '*.xls'))
        file_list += glob.glob(os.path.join(options.search_dir, '*.xlsx'))

    if not file_list:
        warning("BOM file not found..", sys.stdout, terminal=True, prefix="")

        # search version file
        info("Search version file [%s] in [%s]:" % (options.version_file, options.search_dir),\
                sys.stdout, terminal=True, prefix="")
        file_list = glob.glob(os.path.join(options.search_dir, options.version_file))

        # Get bom file list from version.txt
        bom_info = cfg_version(file_list[0])

    if not file_list:
        warning("No version file found..\n", sys.stdout, terminal=True, prefix="")
        parser.print_help()
        sys.exit(1)

    if options.diff and len(file_list) != 2:
        error("In diff mode you should specify only 2 BOMs [%s].\n" % len(file_list), \
                sys.stdout, terminal=True, prefix="")
        parser.print_help()
        sys.exit(1)


    cfg = CfgMergeBom()
    m = MergeBom(file_list, cfg)
    file_list = map(os.path.basename, file_list)

    if options.diff:
        d = m.diff()
        l = m.extra_data()
        write_xls(d, file_list, cfg, options.out_filename, diff=True, extra_data=l)
        sys.exit(0)

    d = m.merge()
    stats = m.statistics()
    write_xls(d, file_list, cfg, options.out_filename, hw_ver=options.rev, \
              pcb_ver=options.pcb_ver, project=options.prj_name, statistics=stats)

