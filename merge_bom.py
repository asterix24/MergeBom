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
import sys, os
import re

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
d = {}
for g in l:
    c = re.search('^[a-zA-Z_]{1,3}', g[REF])
    key = c.group().upper()
    if d.has_key(key):
        d[key].append(g)
    else:
        d[key] = [g]

    print 'Group Key:',c.group(), g[REF]

for k in d.keys():
    d[k] = sorted(d[k], key=lambda ref: ref[REF + 1])

SEPARATOR_NUM = len(l[0]) - 1
ORDER_PATTERN = ['J', 'S','R','C','D','DZ','L', 'Q','TR','Y', 'U']
ORDER_PATTERN_NAMES = {
    'J':['* J Connectors *'],
    'S':['* S Mechanical parts and buttons *'],
    'R':['* R Resistors *'],
    'C':['* C Capacitors *'],
    'D':['* D Diodes *'],
    'DZ':['* DZ Zener, Schottky, Transil *'],
    'L': ['* L Inductors, chokes *'],
    'Q': ['* Q Transistors *'],
    'TR':['* TR Transformers *'],
    'Y': ['* Y Cristal, quarz, oscillators*'],
    'U': ['* U IC *']
}
for i in ORDER_PATTERN_NAMES.keys():
    ORDER_PATTERN_NAMES[i] = (([''] * REF) + ORDER_PATTERN_NAMES[i] + ([''] * (SEPARATOR_NUM - REF)))
    print ORDER_PATTERN_NAMES[i]


for j in d.keys():
    if j not in ORDER_PATTERN:
        print 'Missing order pattern key:'
        print 'In BOM:', d.keys()
        print 'In mergebom:', ORDER_PATTERN
        sys.exit(0)

with open('merged_bom.csv', 'wb') as csvfile:
    data = csv.writer(csvfile, delimiter=';', quotechar='\"')
    data.writerow(header)
    for p in ORDER_PATTERN:
        if d.has_key(p):
            data.writerow(ORDER_PATTERN_NAMES[p])
            for i in d[p]:
                data.writerow(i)

total = 0
for i in table_dict.keys():
    total += table_dict[i][QUANTITY]
    print "Component %s n.%s" % (i, table_dict[i][QUANTITY])

print "Total components %s\n" % total
