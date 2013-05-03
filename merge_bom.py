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
COMMENT=2
FOOTPRINT=3
DESCRIPTION=4

if len(sys.argv) < 2:
    print sys.argv[0], " <csv file name1> <csv file name2> .."
    exit (1)

table = []
CSV_NUM = len(sys.argv[1:])
QUANTITY = CSV_NUM + QUANTITY
REF = CSV_NUM + REF
COMMENT = CSV_NUM + COMMENT

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

        if (j[REF - CSV_NUM][0].lower() == 'j') or (j[DESCRIPTION].lower() == 'test point'):
            key = j[FOOTPRINT] + j[DESCRIPTION]
            print "Except: > ",key
        else:
            key = j[COMMENT - CSV_NUM] + j[FOOTPRINT] + j[DESCRIPTION]

        if key in table_dict:
            table_dict[key][QUANTITY] += int(j[QUANTITY - CSV_NUM])
            table_dict[key][REF] += ", " + j[REF - CSV_NUM]

            if (j[REF - CSV_NUM][0].lower() == 'j') or (j[DESCRIPTION].lower() == 'test point'):
                table_dict[key][COMMENT] += ", " + j[COMMENT - CSV_NUM]
                table_dict[key][index] += int(j[QUANTITY - CSV_NUM])
            else:
                table_dict[key][index] = j[QUANTITY - CSV_NUM]
        else:
            table_dict[key] = pre_col + j
            table_dict[key][QUANTITY] = int(table_dict[key][QUANTITY])
            table_dict[key][index] = int(j[QUANTITY - CSV_NUM])

            if (j[REF - CSV_NUM][0].lower() == 'j') or (j[DESCRIPTION].lower() == 'test point'):
                table_dict[key][COMMENT] = j[COMMENT - CSV_NUM]

    index += 1


l = sorted(table_dict.values(), key=lambda ref: ref[REF][:2])
count = 0
ll = []
for j in l:
    if j[COMMENT].strip() == 'NP':
        ll.append(j)
    else:
        ll.insert(count, j)
        count += 1

with open('merged_bom.csv', 'wb') as csvfile:
    data = csv.writer(csvfile, delimiter=';', quotechar='\"')
    data.writerow(header)
    for i in ll:
        data.writerow(i)

total = 0
for i in table_dict.keys():
    total += table_dict[i][QUANTITY]
    print "Component %s n.%s" % (i, table_dict[i][QUANTITY])

print "Total components %s\n" % total
