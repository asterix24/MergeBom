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

from mergebom_class import *

if __name__ == "__main__":
    import glob
    from optparse import OptionParser

    file_list = []
    parser = OptionParser()
    parser.add_option("-a", "--csv", dest="csv_file", action="store_true",
                      default=False, help="Find and merge CSV files, by defaul are excel files.")
    parser.add_option("-c", "--merge-cfg", dest="merge_cfg",
                      default=None, help="MergeBOM configuration file.")
    parser.add_option("-p", "--search-dir", dest="search_dir",
                      default='./', help="BOM to merge search path.")
    parser.add_option("-o", "--out-filename", dest="out_filename",
                      default='merged_bom.xlsx', help="Out file name")
    parser.add_option("-v", "--version-file", dest="version_file",
                      default='version.txt', help="Version file.")
    parser.add_option(
        "-l",
        "--log-on-file",
        dest="log_on_file",
        default=True,
        action="store_true",
        help="List all project name from version file.")
    parser.add_option(
        "-s",
        "--sub_search_dir",
        dest="sub_search_dir",
        default="Assembly",
        help="Default name for BOM store directory in version file mode.")
    parser.add_option(
        "-d",
        "--diff",
        dest="diff",
        action="store_true",
        default=False,
        help="Generate diff from two specified BOMs")
    parser.add_option(
        "-m",
        "--replace-merged",
        dest="replace_merged",
        action="store_false",
        default=True,
        help="Remove all boms files after merge (only when using version file)")
    parser.add_option("-r", "--bom-revision", dest="bom_rev",
                      default=None, help="Hardware BOM revision")
    parser.add_option("-w", "--bom-pcb-revision", dest="bom_pcb_ver",
                      default=None, help="PCB Revision")
    parser.add_option("-n", "--bom-prj-name", dest="bom_prj_name",
                      default=None, help="Project names.")
    parser.add_option(
        "-t",
        "--bom-date",
        dest="bom_prj_date",
        default=datetime.datetime.today().strftime("%d/%m/%Y"),
        help="Project date.")

    (options, args) = parser.parse_args()

    # Get logger
    logger = report.Report(log_on_file=options.log_on_file, terminal=True)
    logger.write_logo()

    # Load default Configuration for merging
    config = cfg.CfgMergeBom(options.merge_cfg)

    # The user specify file to merge
    file_list = args
    logger.info("Merge BOM file..\n")
    if file_list:
        logger.info("Merge Files:\n")
        logger.info("%s\n" % args)

    if not file_list and os.path.isfile(options.version_file):
        logger.warning("BOM file not found..\n")

        # search version file
        logger.info("Search version file [%s] in [%s]:\n" %
                    (options.version_file,
                     options.search_dir))

        # Get bom file list from version.txt
        """
        Search BOM files in Project directory.
        By defaul the script search in "sub_search_dir" (Assembly) directory all boms like:
        - the folder name contained in "sub_search_dir"/section_name_in_version_file
        - the bom file named as section_name_in_version_file in "sub_search_dir"
        """
        bom_info = cfg.cfg_version(options.version_file)
        for prj in bom_info.keys():
            file_list = []
            glob_path = os.path.join(
                options.search_dir, options.sub_search_dir, prj)

            search_glob = ["*.xls", "*.xlsx"]
            if options.csv_file:
                search_glob = ["*.csv"]

            if not os.path.exists(glob_path):
                glob_path = os.path.join(
                    options.search_dir, options.sub_search_dir)

                search_glob = ["*" + prj + "*.xls", "*" + prj + "*.xlsx"]
                if options.csv_file:
                    search_glob = ["*" + prj + "*.csv"]

            for i in search_glob:
                file_list += glob.glob(os.path.join(glob_path, i))

            if not file_list:
                logger.error("%s: No BOM file to merge\n" % glob_path)
                continue

            m = MergeBom(file_list, config, is_csv=options.csv_file, logger=logger)
            d = m.merge()
            stats = m.statistics()

            tmpdir = tempfile.gettempdir()
            outfilename = os.path.join(
                tmpdir, os.path.basename(
                    options.out_filename))
            dstfilename = os.path.join(
                os.path.dirname(
                    options.out_filename),
                prj +
                "_" +
                os.path.basename(
                    options.out_filename))

            if options.replace_merged:
                outfilename = os.path.join(tmpdir, "bom-%s.xlsx" % prj)
                dstfilename = os.path.join(glob_path, "bom-%s.xlsx" % prj)

            name = bom_info[prj]["name"]
            hw_ver = bom_info[prj]["hw_ver"]
            pcb_ver = bom_info[prj]["pcb_ver"]

            report.write_xls(d,
                             map(os.path.basename, file_list),
                             config,
                             outfilename,
                             hw_ver=hw_ver,
                             pcb_ver=pcb_ver,
                             project=name,
                             statistics=stats,
                             headers= m.header_data())

            if os.path.isfile(outfilename):
                if options.replace_merged:
                    for i in file_list:
                        os.remove(i)

                shutil.copy(outfilename, dstfilename)
                logger.info("Generated %s Merged BOM file.\n" % dstfilename)

        sys.exit(0)

    if not file_list:
        logger.warning(
            "No BOM specified to merge..\n")
        # First merge all xlsx file in search directory
        logger.info("Find in search_dir[%s] all bom files:\n" % options.search_dir)
        file_list = glob.glob(os.path.join(options.search_dir, '*.xls'))
        file_list += glob.glob(os.path.join(options.search_dir, '*.xlsx'))

    # File list empty, so there aren't BOM files to merge, raise error to user.
    if not file_list:
        logger.warning("No version file found..\n")
        parser.print_help()
        sys.exit(1)

    # The user set diff mode, but this works only for two bom.
    if options.diff and len(file_list) != 2:
        logger.error("In diff mode you should specify only 2 BOMs [%s].\n" % len(file_list))
        parser.print_help()
        sys.exit(1)

    m = MergeBom(file_list, config, is_csv=options.csv_file, logger=logger)
    file_list = map(os.path.basename, file_list)

    if options.diff:
        d = m.diff()
        l = m.extra_data()
        report.write_xls(
            d,
            file_list,
            config,
            options.out_filename,
            diff=True,
            extra_data=l,
            headers=m.header_data())
        sys.exit(0)

    d = m.merge()
    stats = m.statistics()

    if options.bom_rev is None \
            or options.bom_pcb_ver is None \
            or options.bom_prj_name is None:
        logger.error("\nYou should specify some missing parameter:\n")
        logger.error("- project name: %s\n" % options.bom_prj_name)
        logger.error("- hw verion: %s\n" % options.bom_rev)
        logger.error("- pcb version: %s\n" % options.bom_pcb_ver)
        sys.exit(1)


    report.write_xls(
        d,
        file_list,
        config,
        options.out_filename,
        hw_ver=options.bom_rev,
        pcb_ver=options.bom_pcb_ver,
        project=options.bom_prj_name,
        statistics=stats,
        headers=m.header_data())

