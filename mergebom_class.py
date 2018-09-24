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

def dump(d):
    for i in d:
        print "Key: %s" % i
        print "Rows [%d]:" % len(d[i])
        for j in d[i]:
            print j
        print "-" * 80


class MergeBom(object):
    def __init__(self, list_bom_files, config, is_csv=False, logger=None, terminal=True):
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
        self.extra_column = []
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
            self.logger.warning("File name %s\n" % file_name)

            # Get all data from select data that could be CSV or xls
            reader = report.DataReader(file_name, is_csv=is_csv)
            data = reader.read()

            # Create filename label for report file
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

                    # Search others marked columns to add to BOM
                    words = re.findall(r'\b\S+\b', item)
                    for w in cfg.VALID_KEYS_CODES:
                        if w in words:
                            item = item.replace(w, '')
                            item = "%s %s" % (w.capitalize(), item.strip())
                            self.extra_column.append((item, n))
                            #print "%s %s" % (item, n)

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
                                row[footprint]
                            ]
                            for ex in self.extra_column:
                                table_dict[r].append(row[ex[1]])

            self.table_list.append(table_dict)

    def table_data(self):
        return self.table_list

    def extra_data(self):
        return self.extra_keys

    def header_data(self):
        return cfg.VALID_KEYS + [k for k, _ in self.extra_column ]

    def group(self):
        self.grouped_items = {}
        for table_dict in self.table_list:
            for designator in table_dict.keys():
                # Group found designator componets by its category
                c = re.search('^[a-zA-Z_]{1,3}', designator)
                group_key = ''
                if c is not None:
                    group_key = self.config.check_category(c.group().upper())

                    if group_key is None:
                        self.logger.error("GROUP key not FOUND!\n")
                        self.logger.error( "%s, %s, %s\n" % (c.group(), designator,
                                                             table_dict[designator]))
                        sys.exit(1)

                    if group_key == '':
                        self.logger.warning("WARNING!! KEY SKIPPED [%s]\n" % group_key)
                        continue

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
                        tmp[key][self.TABLE_TOTALQTY] += item[QUANTITY]
                        tmp[key][curr_file_index] += item[QUANTITY]
                        tmp[key][
                            self.TABLE_DESIGNATOR] += ", " + item[DESIGNATOR]
                        tmp[key][self.TABLE_DESIGNATOR] = lib.order_designator(
                            tmp[key][self.TABLE_DESIGNATOR], self.logger)

                        for ex in self.extra_column:
                            # We add 1 because the table now contain also the
                            # file name, see init function in code that import
                            # data.
                            col_id = ex[1] + 1
                            try:
                                raw_value = item[col_id].strip()
                                if tmp[key][col_id] == "":
                                    tmp[key][col_id] = raw_value
                                else:
                                    if raw_value != "":
                                        words = re.findall(r'\b\S+\b', tmp[key][col_id])
                                        if not raw_value in words:
                                            tmp[key][col_id] += "; " + raw_value

                            except IndexError:
                                print "Error! Impossible to update extra column. This as bug!"
                                sys.exit(1)
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

                        # Add extra column to row if they are present in source
                        # file
                        for ex in self.extra_column:
                            # We add 1 because the table now contain also the
                            # file name, see init function in code that import
                            # data.
                            row.append(item[ex[1] + 1])


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
