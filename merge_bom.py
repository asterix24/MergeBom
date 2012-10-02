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
# Copyright 2012 Daniele Basile <asterix24@gmail.com>
#
#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import sys
import os

QUANTITY=0
REF=1
KEY=2
KEY2=3
KEY3=4

if len(sys.argv) < 3:
    print sys.argv[0], " <csv file name1> <csv file name2> .."
    exit (1)

table = []
CSV_NUM = len(sys.argv[1:])
QUANTITY = CSV_NUM + QUANTITY
REF = CSV_NUM + REF
KEY = CSV_NUM + KEY

header = []
table_dict = {}
pre_col = [0] * CSV_NUM
index = 0

for i in sys.argv[1:]:
    csv_table = csv.reader(open(i, 'rb'), delimiter=',', quotechar='\"')
    for j in csv_table:
        if (j == []):
            continue
        if (j[0] == 'Quantity'):
            header = j
            header = sys.argv[1:] + header
            continue

        key = j[KEY - CSV_NUM] + j[KEY2] + j[KEY3]
        print key
        if key in table_dict:
            table_dict[key][QUANTITY] += int(j[QUANTITY - CSV_NUM])
            table_dict[key][REF] += ", " + j[REF - CSV_NUM]
            table_dict[key][index] = j[QUANTITY - CSV_NUM]
        else:
            table_dict[key] = pre_col + j
            table_dict[key][QUANTITY] = int(table_dict[key][QUANTITY])
            table_dict[key][index] = j[QUANTITY - CSV_NUM]

    index += 1


with open('merged_bom.csv', 'wb') as csvfile:
    data = csv.writer(csvfile, delimiter=',', quotechar='\"')
    data.writerow(header)
    for i in sorted(table_dict.values(), key=lambda ref: ref[REF]):
        data.writerow(i)

total = 0
for i in table_dict.keys():
    total += table_dict[i][QUANTITY]
    print "Component %s n.%s" % (i, table_dict[i][QUANTITY])

print "Total components %s\n" % total
