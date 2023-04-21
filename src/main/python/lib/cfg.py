#!/use/bin/env python
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

"""
MergeBOM Default configuration
"""


import os
import re
import sys
import glob
import configparser

MERGEBOM_VER = "2.0.0"

LOGO_SIMPLE = """

#     #                             ######
##   ## ###### #####   ####  ###### #     #  ####  #    #
# # # # #      #    # #    # #      #     # #    # ##  ##
#  #  # #####  #    # #      #####  ######  #    # # ## #
#     # #      #####  #  ### #      #     # #    # #    #
#     # #      #   #  #    # #      #     # #    # #    #
#     # ###### #    #  ####  ###### ######   ####  #    #
"""

LOGO = """

███╗   ███╗███████╗██████╗  ██████╗ ███████╗██████╗  ██████╗ ███╗   ███╗
████╗ ████║██╔════╝██╔══██╗██╔════╝ ██╔════╝██╔══██╗██╔═══██╗████╗ ████║
██╔████╔██║█████╗  ██████╔╝██║  ███╗█████╗  ██████╔╝██║   ██║██╔████╔██║
██║╚██╔╝██║██╔══╝  ██╔══██╗██║   ██║██╔══╝  ██╔══██╗██║   ██║██║╚██╔╝██║
██║ ╚═╝ ██║███████╗██║  ██║╚██████╔╝███████╗██████╔╝╚██████╔╝██║ ╚═╝ ██║
╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═════╝  ╚═════╝ ╚═╝     ╚═╝
"""

DEFAULT_PRJ_PARAM_KEY = 0
DEFAULT_PRJ_PARAM_LABEL = 1
DEFAULT_PRJ_PARAM_FIELDS = [
    ("prj_date", "Project Date"),
    ("prj_hw_ver", "Hardware Version"),
    ("prj_pcb", "PCB Version"),
    ("prj_name", "Project Name"),
    ("prj_license", "License"),
    ("prj_name_long", "Descriptions"),
    ("prj_part_number", "Part Number"),
    ("prj_status", "Project status"),
    ("prj_prefix", "Project group")
]

DEFAULT_PRJ_PARAM_DICT = {
    "prj_date": "--/--/----",
    "prj_hw_ver": "0",
    "prj_pcb": "A",
    "prj_name": "MyProject",
    "prj_license": "Copyright to me",
    "prj_name_long": "My usefull project",
    "prj_part_number": "--",
    "prj_status": "Prototype",
    "prj_prefix": "--",
}

# pcb-{prefix}{subprjname}-Rev-{PCB-ver}
TEMPLATE_PCB_NAME_PANEL = "pcb-%s%s_panel-Rev-%s"
# pcb-{prefix}{subprjname}-Rev-{PCB-ver}
TEMPLATE_PCB_NAME_PANNEL = "pcb-%s%s_pannel-Rev-%s"
TEMPLATE_PCB_NAME = "pcb-%s%s-Rev-%s"  # pcb-{prefix}{subprjname}-Rev-{PCB-ver}
TEMPLATE_PRJ_NAME = "%s%s-Rev.%s"        # {prefix}{subprjname}-Rev.{hw-ver}
TEMPLATE_HW_DIR = "Hw-Rev.%s%s"
VERSION_FILE = "version.txt"

DEFAULT_PRJ_DIR = [
    ("Pdf", 'schematic*.pdf'),
    ("Assembly", 'bom*.xlsx'),
    ("Assembly", 'assembly*.pdf'),
    ("Assembly", 'pick-and-place*.txt'),
]

FILE_TO_SKIP = [
    "Status File.txt"
]

ENG_LETTER = {
    'G': (1e9, 1e8),
    'M': (1e6, 1e5),
    'k': (1e3, 1e2),
    'R': (1, 0.1),
    'm': (1e-3, 1e-4),
    'u': (1e-6, 1e-7),
    'n': (1e-9, 1e-10),
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

VALID_KEYS_CODES = [
    u'Layer',
    u'MountTechnology',
    u'Mounting_Technology',
    u'CODE',
    u'NOTE',
    u'NP_Value',
]

EXTRA_KEYS = [
    u'date',
    u'project',
    u'hardware_version',
    u'pcb_version',
]

CATEGORY_NAMES_DEFAULT = [
    {
        'name': 'Labeling',
        'desc': 'Labeling and Info',
        'group': [],
        'ref': 'LABEL',
    },
    {
        'name': 'Connectors',
        'desc': 'Connectors and holders',
        'group': ['X', 'P', 'SIM', 'CN', 'JP', 'XJ'],
        'ref': 'J',
    },
    {
        'name': 'Mechanicals',
        'desc': 'Mechanical parts and buttons',
        'group': [
            'SCR',
            'SPA',
            # Battery
            'SCREW',
            'WASHER',
            'NUT',
            'MEC',
            'BAT',
            'BUZ',  # Buzzer
            # Buttons
            'BT',
            'B',
            'SW',
            'MP',  # spacer and stud
            'K', 'M'],
        'ref': 'S',
    },
    {
        'name': 'Fuses',
        'desc': 'Fuses discrete components, varistor',
        'group': ['FU', 'VR'],
        'ref': 'F',
    },
    {
        'name': 'Resistors',
        'desc': 'Resistor components',
        'group': ['RN', 'R_G'],
        'ref': 'R',
    },
    {
        'name': 'Relays',
        'desc': 'Relays components',
        'group': [''],
        'ref': 'RL',
    },
    {
        'name': 'Capacitors',
        'desc': 'Capacitors',
        'group': [],
        'ref': 'C',
    },
    {
        'name': 'Diode',
        'desc': 'Diodes, Zener, Schottky, LED, Transil',
        'group': ['DZ', 'FC'],
        'ref': 'D',
    },
    {
        'name': 'Inductors',
        'desc': 'L  Inductors, chokes, Ferrite',
        'group': ['FB'],
        'ref': 'L',
    },
    {
        'name': 'Transistor',
        'desc': 'Q Transistors, MOSFET, Triac',
        'group': ['TRC'],
        'ref': 'Q',
    },
    {
        'name': 'Transformes',
        'desc': 'TR Transformers',
        'group': ['T'],
        'ref': 'TR',
    },
    {
        'name': 'Cristal',
        'desc': 'Cristal, quarz, oscillator',
        'group': [],
        'ref': 'Y',
    },
    {
        'name': 'IC',
        'desc': 'Integrates and chips',
        'group': [],
        'ref': 'U',
    },
    {
        'name': 'DISCARD',
        'desc': 'Reference to discard, to not put in BOM',
        'group': ['TP'],
        'ref': '',
    },
]


NOT_POPULATE_KEY = ["NP", "NM"]
NP_REGEXP = r"^NP\s"

MERGED_FILE_TEMPLATE = "bom-%s_R%s.xlsx"
MERGED_FILE_TEMPLATE_VARIANT = "bom-%s_var-%s_R%s.xlsx"
MERGED_FILE_TEMPLATE_HW = "%s%s_R%s.xlsx"
MERGED_FILE_TEMPLATE_NOHW = "%s%s_merged.xlsx"

PRJ_DATE = 'prj_date'
PRJ_HW_VER = 'prj_hw_ver'
PRJ_LICENSE = 'prj_license'
PRJ_NAME = 'prj_name'
PRJ_NAME_LONG = 'prj_name_long'
PRJ_PCB = 'prj_pcb'
prj_part_number = 'prj_part_number'
PRJ_STATUS = 'prj_status'


class CfgMergeBom(object):
    """
    MergeBOM Configuration
    """

    def __init__(self, cfgfile_name=None, handler=sys.stdout, terminal=True, logger=None):
        self.handler = handler
        self.terminal = terminal
        self.category_names = CATEGORY_NAMES_DEFAULT

        if cfgfile_name is not None:
            try:
                config_file = open(cfgfile_name)
                config = toml.loads(config_file.read())
                self.category_names = config.get('category_names', None)
            except IOError as e:
                logger.error("Configuration: %s" % e,
                             self.handler, terminal=self.terminal)
                logger.warning("No Valid Configuration file! Use Default",
                               self.handler, terminal=self.terminal)

        if self.category_names is None:
            logger.warning("No Valid Configuration file! Use Default",
                           self.handler, terminal=self.terminal)

    def check_category(self, group_key):
        if not group_key:
            return group_key

        for item in self.category_names:
            if group_key in item['group'] or group_key == item['ref']:
                return item['ref']

        return None

    def categories(self):
        categories = []
        for item in self.category_names:
            if item['ref']:
                categories.append(item['ref'])

        return categories

    def get(self, category, key):
        """
        By category and key get all related info
        """
        for item in self.category_names:
            if item['ref'] == category:
                return item[key]

        return None


def extrac_projects(wk_file):
    """
    [ProjectGroup]
    Version=1.0
    [Project1]
    ProjectPath=camera.PrjMbd
    [Project2]
    ProjectPath=5mp-sensor\5mp-sensor.PrjPCB
    [Project3]
    ProjectPath=camera-core\camera-core.PrjPCB
    [Project4]
    ProjectPath=18mp-sensor\18mp-sensor.PrjPCB
    [Project5]
    ProjectPath=imx8m_evk\imx8_evk.PrjPcb
    """
    wk_config = configparser.RawConfigParser()
    wk_config.read(wk_file)
    l = []
    for i in wk_config.sections():
        try:
            if re.match("project[0-9]+", i.lower()) is None:
                continue
            temp = wk_config.get(i, 'ProjectPath')
            if "prjpcb" not in temp.lower():
                continue

            if "\\" in temp:
                temp = temp.split('\\')
                basename = os.path.join(*temp[:-1])
                complete_path = os.path.join(*temp)
                l.append((basename, complete_path))
            else:
                if "." in temp:
                    basename = temp.split(".")[0]
                complete_path = temp
                l.append((basename, complete_path))

        except configparser.NoOptionError:
            continue

    return l


import io


def get_parameterFromPrj(prj_name, prj_file):
    """
    Get paramet from Altium project.
    """
    d = {}
    if not os.path.exists(prj_file):
        return "", {}

    prj_config = configparser.RawConfigParser()
    with io.open(prj_file, 'r', encoding='utf_8_sig') as fp:
        prj_config.readfp(fp)

    for i in prj_config.sections():
        if re.match(r'Parameter[0-9]+', i) is None:
            continue
        parametro = prj_config.get(i, 'Name')
        val = prj_config.get(i, 'Value')
        d[parametro] = val

    return prj_name, d


def get_variantFromPrj(prj_name, prj_file):
    """
    Get paramet from Altium project.
    """
    variants = []
    if not os.path.exists(prj_file):
        return []

    prj_config = configparser.RawConfigParser()
    with io.open(prj_file, 'r', encoding='utf_8_sig') as fp:
        prj_config.readfp(fp)

    for i in prj_config.sections():
        if re.match(r'ProjectVariant[0-9]+', i) is None:
            continue
        variants.append(prj_config.get(i, 'Description'))

    return variants


def find_bomfiles(root_path, prj_name, csv_file):
    """
    find all bom files in prj
    """
    flt = "*.xlsx"
    if csv_file:
        flt = "*.csv"

    pth = os.path.join(root_path, prj_name, flt)
    bom_list = glob.glob(pth)
    return [prj_name, bom_list]


def cfg_altiumWorkspace(workspace_file_path, csv_file, bom_search_dir,
                        logger, bom_postfix="", bom_prefix="bom-"):
    """
    Alla funzione vengono passati i parametri:
        1. il path del file Workspace
        2. se i file da mergiare sono di tipo csv o xlsx
        3. nome del file con cui fare il merge ricerca del nome di tutti
        i progetti all'interno del file Workspace
    esempio di file Wprkspace:
        $ cat schemes.DsnWrk
        [ProjectGroup]
        Version=1.0
        [Project1]
        ProjectPath=camera-tbd\camera-tbd.PrjPcb
        [Project2]
        ProjectPath=usb-serial\\usb-serial.PrjPcb
    """

    ret = []

    # calcolo path dove si trovano i progetti
    root_path = os.path.dirname(workspace_file_path)
    file_to_merge_path = os.path.join(root_path, bom_search_dir)

    logger.info("\nSearch project to merge in given Altiumworkspace: %s\n" %
                workspace_file_path)
    logger.info("BOM path %s\n" % file_to_merge_path)
    logger.info("Root path %s\n" % root_path)

    wk_config = configparser.RawConfigParser()
    wk_config.read(workspace_file_path)
    for i in wk_config.sections():
        try:
            if re.match("project[0-9]+", i.lower()) is None:
                #logger.error("Skip key %s\n" % i)
                continue

            temp = wk_config.get(i, 'ProjectPath')
            if "\\" in temp:
                temp = temp.split('\\')
                basename = os.path.join(*temp[:-1])
                complete_path = os.path.join(*temp)
            else:
                if "." in temp:
                    basename = temp.split(".")[0]
                complete_path = temp

        except configparser.NoOptionError:
            #logger.info("Missing key %s\n" % i)
            continue

        parametri_dict = {}
        file_BOM = []

        # ricerca parametri di ogni progetto e poi messi in un dizionario con
        # {nomeparametro : parametro}
        prj = os.path.join(root_path, complete_path)
        if not os.path.exists(prj):
            logger.error("Unable to find project BOM: %s\n" % prj)
            continue

        prj_config = configparser.RawConfigParser()
        prj_config.read(prj)

        logger.info("proj %s\n" % prj)
        for i in prj_config.sections():
            if re.match(r'Parameter[0-9]+', i) is None:
                #logger.info("Wrong key found [%s]\n" % i)
                continue

            parametro = prj_config.get(i, 'Name')
            val = prj_config.get(i, 'Value')
            parametri_dict[parametro] = val

        # ricerca file del progetto a cui fare il merge e messi in una lista
        bom_name = "%s%s%s" % (bom_prefix, basename, bom_postfix)
        path_file = os.path.join(file_to_merge_path, basename)
        merge_file_item = os.path.join(path_file, bom_name) + '.csv'
        if not csv_file:
            merge_file_item = os.path.join(path_file, bom_name) + '.xlsx'

        if os.path.exists(merge_file_item):
            file_BOM.append(merge_file_item)

        # creo una tupla con il dizionario dei parametri e la lista dei file e lo metto all'interno di un'altra lista (ret):
        # [
        #   ([file1.csv, file2.csv], {nomeparametro : parametro})
        #   ([file1.csv, file2.csv], {nomeparametro : parametro})
        # ]
        if not (parametri_dict == {} and file_BOM == []):
            ret.append((basename, file_BOM, parametri_dict))

    return ret


if __name__ == "__main__":
    import toml

    if len(sys.argv) < 2:
        print("Usage %s <cfg filename>" % sys.argv[0])
        sys.exit(1)

    config = "Vuoto"
    with open(sys.argv[1]) as configfile:
        config = toml.loads(configfile.read())

    print(type(config), len(config))
    print(config.keys())
    print(config['category_names'])
