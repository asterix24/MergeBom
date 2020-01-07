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

import re
from src.main.python.lib.cfg import CATEGORY_TO_UNIT, NOT_POPULATE_KEY, ENG_LETTER

def order_designator(ref_str, logger):
    ref_str = ref_str.replace(" ", "")
    l = ref_str.split(",")
    try:
        d = sorted(l, key=lambda x: int(re.search('[0-9]+', x).group()))
    except TypeError:
        s = "Could not order Designators [%s]\n" % l
        logger.error(s)
        raise Exception(s)

    return ", ".join(d)

def value_toFloat(l, unit, logger):
    acc = 0
    value = "0"
    mult = 1
    div = 1
    if unit not in CATEGORY_TO_UNIT:
        s = "Unknow category [%s] allowed are[%s]\n" % (unit, CATEGORY_TO_UNIT.keys())
        logger.error(s)
        raise Exception(s)

    # K is always chilo .. so fix case
    l = l.replace("K", "k")
    # herz
    l = l.replace("HZ", "Hz")
    l = l.replace("hz", "Hz")
    l = l.replace("ohm", "R")

    # manage correctly NP value
    for n in NOT_POPULATE_KEY:
        if n in l:
            return -1, l, ""

    # In string value we could find a note or other info, remove it, use later
    note = ""
    if " " in l:
        ss = l.split(" ")
        l = ss[0]
        note = ss[1:]
        # print ">", l, note

    # First character should be a numer
    if re.search("^[0-9]", l) is None:
        return -2, l, note

    cnt = 0
    flag = False
    for c in l:
        if c in ENG_LETTER and not flag:
            try:
                value = value.replace(',', '.')
                acc = float(value)
                mult, div = ENG_LETTER[c]
                value = "0"
                flag = True
                continue
            except ValueError as e:
                s = "l[%s] Acc[%s], mult[%s], value[%s], div[%s], {%s}\n" % (l, acc, mult, value, div, e)
                logger.error(s)
                raise Exception(s)

        # skip measure unit, we use only numbers
        if re.search("[a-zA-z]", c) is not None:
            continue

        value += c

        if flag:
            cnt += 1

    try:
        div2 = 1
        for c in range(0, cnt-1):
            div2 = div2 * 10

        value = acc * mult + float(value) * (div / div2)
    except ValueError as e:
        logger.error(">> l[%s] Acc[%s], mult[%s], value[%s], div[%s], {%s}\n" % (
            l, acc, mult, value, div, e))
        return -2, l, note

    return value, CATEGORY_TO_UNIT[unit], note

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
    x3 = x / math.pow(10, exp3)
    exp3_text = ""

    if -24 <= exp3 <= 24 and exp3 != 0:
        exp3_text = 'yzafpnum kMGTPEZY'[int((exp3 - (-24)) / 3)]

    return sign, str(x3), exp3_text


def value_toStr(l, logger):
    try:
        value, unit, note = l
    except ValueError as e:
        s = "Unpack error %s {%s}\n" % (l, e)
        logger.error(s)
        raise Exception(s)

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
        if unit in ['Hz']:
            number = number.rstrip(".0")
            number += notation
        else:
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

