
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
# Copyright 2018 Daniele Basile <asterix24@gmail.com>
#

#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import os
import glob
import argparse
import ConfigParser
import re

def calcProjects(projectpath):
    valuesDict = {}

    if len(glob.glob(os.path.join(projectpath, '*.DsnWrk'))) == 1:
        work = glob.glob(os.path.join(projectpath, '*.DsnWrk'))

        config = ConfigParser.RawConfigParser()
        config.read(work)

        for i in config.sections():
            try:
                temp = config.get(i, 'ProjectPath')
                temp = temp.split('/')
                k = []
                p = []

                if len(temp) > 1:
                    k = os.path.splitext(temp[1])[0]
                    p = os.path.join(*temp)
                else:
                    k = os.path.splitext(temp[0])[0]
                    p = os.path.join("./", *temp)
                    
                if len(k) == 0 or len(p) == 0:
                    print ("Errore nel parsing del workspace")
                    sys.exit(1)

                k = k.lower()
                valuesDict[k] = p

            except ConfigParser.NoOptionError:
                pass

        return True, valuesDict
    else:
        print ('Nessun file DsnWrk o piu di 1 file DsnWrk')
        return False, 'None'

    
def ricercaParametri(file, pathproject):
    parametri_dict={}

    parametri=['prj_date','prj_hw_ver','prj_license', 'prj_name', 'prj_name_long', 'prj_pcb', 'prj_pn', 'prj_status']

    prj=os.path.join(pathproject,file)
    f=open(prj,'r')
    config=ConfigParser.RawConfigParser()
    config.read(prj)

    for i in config.sections():
        line = re.findall(r'Parameter[0-9]', i)

        if line:
            parametro=config.get(i,'Name')
            val=config.get(i,'Value')
            parametri_dict[parametro]=val

    return parametri_dict


if __name__ == "__main__":
    

    file_list = []
    parser = argparse.ArgumentParser()
    parser.add_argument('--workspace_file', '-w', dest='ws', 
                        help='Dove si trova il file WorkSpace', default='./test')
    parser.add_argument("-a", "--csv", dest="csv_file", action="store_true",
                      default=False, help="Find and merge CSV files, by defaul are excel files.")
    options=parser.parse_args()

    trovato,path_dict=calcProjects(options.ws)    #se file Wrk trovato allora trovato=True    path_dict=dizionario con tutti i progetti
    print(path_dict)
    if trovato:
        nProgetti=0 #ricerca parametri per ogni progetto e creazione dizionario con {nome progetto: {parametro: valore}}
        progetti_dict={}
        for k, v in path_dict.items():
            parametri_dict={}
            parametri_dict=ricercaParametri(v, options.ws)
            progetti_dict[k]=parametri_dict
            nProgetti+=1
        print(progetti_dict)

        filemerge=os.path.join(options.ws, "Assembly")

    else:
        print("file WorkSpace non trovato!")
