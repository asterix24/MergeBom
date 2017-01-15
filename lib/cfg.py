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
# Copyright 2017 Daniele Basile <asterix24@gmail.com>
#

MERGEBOM_VER="1.0.0"

logo_simple = """

#     #                             ######
##   ## ###### #####   ####  ###### #     #  ####  #    #
# # # # #      #    # #    # #      #     # #    # ##  ##
#  #  # #####  #    # #      #####  ######  #    # # ## #
#     # #      #####  #  ### #      #     # #    # #    #
#     # #      #   #  #    # #      #     # #    # #    #
#     # ###### #    #  ####  ###### ######   ####  #    #

"""

logo = """
███╗   ███╗███████╗██████╗  ██████╗ ███████╗██████╗  ██████╗ ███╗   ███╗
████╗ ████║██╔════╝██╔══██╗██╔════╝ ██╔════╝██╔══██╗██╔═══██╗████╗ ████║
██╔████╔██║█████╗  ██████╔╝██║  ███╗█████╗  ██████╔╝██║   ██║██╔████╔██║
██║╚██╔╝██║██╔══╝  ██╔══██╗██║   ██║██╔══╝  ██╔══██╗██║   ██║██║╚██╔╝██║
██║ ╚═╝ ██║███████╗██║  ██║╚██████╔╝███████╗██████╔╝╚██████╔╝██║ ╚═╝ ██║
╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═════╝  ╚═════╝ ╚═╝     ╚═╝

"""

ENG_LETTER = {
    'G': (1e9,     1e8),
    'M': (1e6,     1e5),
    'k': (1e3,     1e2),
    'R': (1,       0.1),
    'u': (1e-6,   1e-7),
    'n': (1e-9,  1e-10),
    'p': (1e-12, 1e-13),
}

CATEGORY_TO_UNIT = {
    'R': "ohm",
    'C': "F",
    'L': "H",
    'Y': "Hz",
}

VALID_KEYS = [
    u'designator',
    u'comment',
    u'footprint',
    u'description',
]

EXTRA_KEYS = [
    u'date',
    u'project',
    u'hardware_version',
    u'pcb_version',
]

CONNECTORS   = 'J'
MECHANICALS  = 'S'
FUSES        = 'F'
RESISTORS    = 'R'
CAPACITORS   = 'C'
DIODES       = 'D'
ZENERS       = 'DZ'
INDUCTORS    = 'L'
TRANSISTORS  = 'Q'
TRANSFORMERS = 'TR'
CRYSTALS     = 'Y'
INTEGRATES   = 'U'

CATEGORY_NAMES = {
    CONNECTORS   :'J  Connectors',
    MECHANICALS  :'S  Mechanical parts and buttons',
    FUSES        :'F  Fuses',
    RESISTORS    :'R  Resistors',
    CAPACITORS   :'C  Capacitors',
    DIODES       :'D  Diodes',
    ZENERS       :'DZ Zener, Schottky, Transil',
    INDUCTORS    :'L  Inductors, chokes',
    TRANSISTORS  :'Q  Transistors',
    TRANSFORMERS :'TR Transformers',
    CRYSTALS     :'Y  Cristal, quarz, oscillator',
    INTEGRATES   :'U  IC',
}


NOT_POPULATE_KEY = ["NP", "NM"]
NP_REGEXP = r"^NP\s"
