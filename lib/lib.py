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

import cfg

def order_designator(ref_str):
    ref_str = ref_str.replace(" ", "")
    l = ref_str.split(",")
    try:
        d = sorted(l, key=lambda x: int(re.search('[0-9]+', x).group()))
    except TypeError:
        error("Could not order Designators [%s]" % l)
        sys.exit(1)
    return ", ".join(d)


def value_toFloat(l, unit, handler=sys.stdout, terminal=True):
    acc = 0
    value = "0"
    mult = 1
    div = 1
    if unit not in cfg.CATEGORY_TO_UNIT:
        error(
            "Unknow category [%s] allowed are[%s]" %
            (unit, cfg.CATEGORY_TO_UNIT.keys()), handler, terminal=terminal)
        sys.exit(1)

    # K is always chilo .. so fix case
    l = l.replace("K", "k")

    # manage correctly NP value
    for n in cfg.NOT_POPULATE_KEY:
        if n in l:
            return -1, l, ""

    # In string value we could find a note or other info, remove it, put for
    # later
    note = ""
    if " " in l:
        ss = l.split(" ")
        l = ss[0]
        note = ss[1:]
        # print ">", l, note

    # First character should be a numer
    if re.search("^[0-9]", l) is None:
        return -2, l, note

    for c in l:
        if c in cfg.ENG_LETTER:
            try:
                value = value.replace(',', '.')
                acc = float(value)
                mult, div = cfg.ENG_LETTER[c]
                value = "0"
                continue
            except ValueError as e:
                error(
                    "l[%s] Acc[%s], mult[%s], value[%s], div[%s], {%s}" %
                    (l, acc, mult, value, div, e), handler, terminal=terminal)
                sys.exit(1)

        if c in cfg.CATEGORY_TO_UNIT[unit]:
            continue

        value += c
        # print "[",c,"<>", value, "]",

    try:
        value = acc * mult + float(value) * div
    except ValueError as e:
        error("l[%s] Acc[%s], mult[%s], value[%s], div[%s], {%s}" % (
            l, acc, mult, value, div, e), handler, terminal=terminal)
        return -2, l, note

    return value, cfg.CATEGORY_TO_UNIT[unit], note

import math


def eng_string(x):
    '''
    Returns float/int value <x> formatted in a simplified engineering format -
    using an exponent that is a multiple of 3.

    format: printf-style string used to format the value before the exponent.
          1230.0 => 1.23k
      -1230000.0 => -1.23M
    '''
    x = float(x)
    sign = ''
    if x < 0:
        x = -x
        sign = '-'
    exp = int(math.floor(math.log10(x)))
    exp3 = exp - (exp % 3)
    #x3 = x / (10 ** exp3)
    x3 = x / math.pow(10, exp3)

    if exp3 >= -24 and exp3 <= 24 and exp3 != 0:
        exp3_text = 'yzafpnum kMGTPEZY'[(exp3 - (-24)) / 3]
    elif exp3 == 0:
        exp3_text = ''

    return (sign, str(x3), exp3_text)


def value_toStr(l, handler=sys.stdout, terminal=True):
    try:
        value, unit, note = l
    except ValueError as e:
        error("Unpack error %s {%s}" % (l, e), handler, terminal=terminal)
        sys.exit(1)

    if value in [-1, -2]:
        return "%s %s" % (unit, " ".join(note))

    if value == 0.0:
        sign, number, notation = "", "0.0", ""
    else:
        sign, number, notation = eng_string(value)

    if notation == "":
        if unit == "ohm":
            number = str(value)
            unit = "R"
        number = re.sub(r"\.0+$", '', number)

    elif notation in ["k", "M", "G", "T", "P", "E", "Z", "Y"]:
        number = number.replace(".", notation)
        number = number.rstrip("0")

        if unit == "ohm":
            unit = ""

    elif notation in ["y", "z", "a", "f", "p", "n", "u", "m"]:
        number = number.rstrip("0")
        if notation == "m" and unit == "ohm":
            number = str(value)
            unit = "R"
        else:
            number = re.sub(r"\.$", "", number)
            number = "%s%s" % (number, notation)

    space = ""
    if note:
        space = " "
    return "%s%s%s%s%s" % (sign, number, unit, space, " ".join(note))

