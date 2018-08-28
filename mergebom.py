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
import glob
import argparse
import ConfigParser
import re
from lib import cfg
   
if __name__ == "__main__":

    file_list = []
    parser = argparse.ArgumentParser()
    parser.add_argument('--workspace_file', '-w', dest='ws', 
                        help='Dove si trova il file WorkSpace', default='./test')
    parser.add_argument("-a", "--csv", dest="csv_file", action="store_true",
                      default=False, help="Find and merge csv files, by defaul are excel files.")
    options=parser.parse_args()
    file_BOM={}
    file_BOM, progetti_dict=cfg.cfg_altiumWorkspace(options)
    if not file_BOM:
        print("i file non sono stati trovati")

