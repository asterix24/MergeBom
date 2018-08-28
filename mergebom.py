
#!/usr/bin/env python
# -*- coding: utf-8 -*-

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

    trovato,path_dict=cfg.calc_projects(options.ws)    #se file Wrk trovato allora trovato=True    path_dict=dizionario con tutti i progetti
    print(path_dict)
    if trovato:
        file_BOM, progetti_dict=cfg.ricerca_e_verifica(options, path_dict)
        if not file_BOM:
            print("i file non sono stati trovati")

    else:
        print("file WorkSpace non presente nella directori indicata")