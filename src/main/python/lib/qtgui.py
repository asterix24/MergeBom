#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import shutil
from pathlib import Path

from src.main.python.lib.cfg import LOGO, MERGEBOM_VER
from src.main.python.lib.cfg import extrac_projects, get_parameterFromPrj, find_bomfiles, get_variantFromPrj
from src.main.python.lib.cfg import MERGED_FILE_TEMPLATE, LOGO_SIMPLE, DEFAULT_PRJ_PARAM_DICT, MERGED_FILE_TEMPLATE_VARIANT
from src.main.python.lib.cfg import TEMPLATE_PCB_NAME, TEMPLATE_PRJ_NAME, TEMPLATE_HW_DIR
from src.main.python.lib.cfg import VERSION_FILE, DEFAULT_PRJ_DIR, DEFAULT_PRJ_DIR_VARIANT
from src.main.python.lib.cfg import CfgMergeBom
from src.main.python.lib.common import copyGerberZip
from src.main.python.lib.report import ReportBase, write_xls
from src.main.python.mergebom_class import *
import src.main.python.resources

from PyQt5.QtCore import QDateTime, Qt, pyqtSlot
from PyQt5.QtWidgets import (QApplication, QTextEdit, QCheckBox, QDialog, QGroupBox,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, QStyleFactory,
                             QTableWidget, QVBoxLayout, QWidget, QFileDialog, QListWidget,
                             QTableWidgetItem, QHeaderView, QComboBox)

from PyQt5.QtGui import QIcon, QFont, QFontDatabase

SUPPORTED_FILE_TYPE = [
    "Altium WorkSpace (*.DsnWrk)",
    "Altium Project (*.PrjPCB)",
    "BOM files (*.xlsx *.xls *.csv)",
]

NO_VARIANT = "[No variant]"


class Report(ReportBase):
    """
    Report class
    """

    def __init__(self, logWidget):
        super(Report, self).__init__()
        self.logWidget = logWidget

    def printout(self, s, prefix="", color=""):
        s = s.replace('\n', "<br>")
        line = "%s%s" % (prefix, s)
        if color != "":
            line = "<font color=\"%s\">%s%s</font>" % (color, prefix, s)
        self.logWidget.insertHtml(line)
        self.logWidget.verticalScrollBar().setValue(
            self.logWidget.verticalScrollBar().maximum())

    def write_logo(self):
        self.printout(LOGO_SIMPLE, color='green')


class MergeBomGUI(QDialog):
    """
    GUI frontend of mergebom script.
    With the GUI is possible also to deploy the proct on given directory.
    """

    def __init__(self, parent=None):
        super(MergeBomGUI, self).__init__(parent)

        self.tmp_bom_list = []
        self.prj_and_data = {}
        self.param_prj_list_view = None
        self.param_bom_list_view = None

        QFontDatabase.addApplicationFont(":/fonts/MesloLGS NF Regular.ttf")
        self.main_layout = QVBoxLayout()
        self.sub_layout = QHBoxLayout()
        self.q1_layout = QVBoxLayout()
        self.q2_layout = QVBoxLayout()
        self.sub_top_q2_layout = QHBoxLayout()
        self.sub_layout.addLayout(self.q1_layout, 10)
        self.sub_layout.addLayout(self.q2_layout, 90)

        # Q1 right quadrant
        self.merge_sel = QPushButton("Merge Select")
        self.merge_sel.setDefault(True)
        self.merge_sel.clicked.connect(self.__merge_bom_select)
        self.merge_all = QPushButton("Merge All")
        self.merge_all.setDefault(True)
        self.merge_all.clicked.connect(self.__merge_bom_all)

        self.merge_only_csv = QCheckBox("CSV Filter", self)
        self.merge_only_csv.setChecked(True)
        self.merge_only_csv.stateChanged.connect(self.__filter_checkbox)
        self.merge_prj_files = QCheckBox("Only Prj BOM", self)
        self.merge_prj_files.setChecked(True)
        self.merge_prj_files.stateChanged.connect(self.__filter_prj)
        self.merge_same_dir = QCheckBox("Merge in same directory", self)
        self.merge_same_dir.stateChanged.connect(self.__check_same_dir)
        self.merge_same_dir_path = QLineEdit(os.path.expanduser('~/'))
        self.merge_same_dir_path_select = QPushButton("Select")
        self.merge_same_dir_path_select.clicked.connect(
            self.__select_merge_path)
        self.delete_merged = QCheckBox("Delete Mergeded", self)
        self.merge_bom_outname_label = QLabel("Out Merged file name:")
        self.merge_bom_outname = QLineEdit("merged_bom.xlsx")
        self.merge_bom_outname.setEnabled(False)
        self.merge_autoname = QCheckBox("Autoname", self)
        self.merge_autoname.setChecked(True)
        self.merge_autoname.stateChanged.connect(self.__autoname_out_file)

        self.merge_cmd_box = QGroupBox("MergeBOM Commands")
        self.merge_cmd_box.setEnabled(False)
        vbox = QVBoxLayout()
        self.merge_cmd_box.setLayout(vbox)
        vbox.addWidget(self.merge_sel)
        vbox.addWidget(self.merge_all)
        vbox.addWidget(self.merge_only_csv)
        vbox.addWidget(self.merge_prj_files)
        vbox.addWidget(self.merge_same_dir)
        vbox.addWidget(self.merge_same_dir_path)
        vbox.addWidget(self.merge_same_dir_path_select)
        vbox.addWidget(self.delete_merged)
        vbox.addWidget(self.merge_bom_outname_label)
        vbox.addWidget(self.merge_bom_outname)
        vbox.addWidget(self.merge_autoname)

        self.deploy_sel = QPushButton("Deploy Selected Prj")
        self.deploy_sel.setDefault(True)
        self.deploy_sel.clicked.connect(self.__deploy_prj_select)
        self.deploy_all = QPushButton("Deploy All Prj")
        self.deploy_all.setDefault(True)
        self.deploy_all.clicked.connect(self.__deploy_all_prj)
        self.deploy_path_label = QLabel("Deploy Path:")
        self.deploy_path = QLineEdit(os.path.join(
            os.path.expanduser('~/'), "Dropbox", "Produzione", "FileProgetto"))
        self.deploy_path_select = QPushButton("Select")
        self.deploy_path_select.clicked.connect(self.__select_deploy_path)
        self.deploy_customer_name_label = QLabel("Customer Name:")
        self.deploy_customer_name = QLineEdit("Customer")
        self.deploy_rev_name_label = QLabel("Revision:")
        self.deploy_rev_name = QLineEdit("0")

        self.deploy_cmd_box = QGroupBox("Deploy Commands")
        self.deploy_cmd_box.setEnabled(False)
        vbox = QVBoxLayout()
        self.deploy_cmd_box.setLayout(vbox)
        vbox.addWidget(self.deploy_sel)
        vbox.addWidget(self.deploy_all)
        vbox.addWidget(self.deploy_path_label)
        vbox.addWidget(self.deploy_path)
        vbox.addWidget(self.deploy_path_select)
        vbox.addWidget(self.deploy_customer_name_label)
        vbox.addWidget(self.deploy_customer_name)
        vbox.addWidget(self.deploy_rev_name_label)
        vbox.addWidget(self.deploy_rev_name)

        self.q1_layout.addWidget(self.deploy_cmd_box)
        self.q1_layout.addWidget(self.merge_cmd_box)
        self.q1_layout.addStretch(1)

        # Q2 left quadrant

        # Altium workspace selection
        self.selected_file = QLineEdit(os.path.expanduser('~/'))
        btn = QPushButton("Select")
        btn.setDefault(True)
        btn.clicked.connect(self.__select_wk_file)
        self.label_param = QLabel(
            "Select Altium workspace, project or BOM files: --")
        top_hbox = QHBoxLayout()
        top_hbox.addWidget(self.selected_file)
        top_hbox.addWidget(btn)

        self.param_prj_list_view = QListWidget()
        self.param_prj_list_view.itemClicked.connect(self.__on_click_list)
        self.param_prj_list_view.currentTextChanged.connect(
            self.__on_click_list)
        self.param_bom_list_view = QListWidget()

        self.param_table_view = QTableWidget()
        hdr = self.param_table_view.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Stretch)
        self.param_table_view.setRowCount(10)
        self.param_table_view.setColumnCount(2)
        self.param_table_view.setHorizontalHeaderLabels(["Param", "Value"])

        self.log_panel = QTextEdit(self)
        self.log_panel.setReadOnly(True)
        self.log_panel.setFont(QFont("MesloLGS NF", 10))

        sub_vbox = QVBoxLayout()
        sub_vbox.addWidget(QLabel("Projects:"))
        sub_vbox.addWidget(self.param_prj_list_view)
        sub_vbox2 = QVBoxLayout()
        sub_vbox2.addWidget(QLabel("Merge Param:"))
        sub_vbox2.addWidget(self.param_table_view)
        self.sub_top_q2_layout.addLayout(sub_vbox, 40)
        self.sub_top_q2_layout.addLayout(sub_vbox2, 60)

        self.bom_list_search_path = QLineEdit()
        self.bom_list_search_select = QPushButton("Search Path")
        self.bom_list_search_select.setDefault(True)
        self.bom_list_search_select.clicked.connect(
            self.__select_bom_search_path)
        self.prj_variant = QComboBox()
        self.prj_variant.activated.connect(self.__select_prj_variant)

        sub_vbox3 = QHBoxLayout()
        sub_vbox3.addWidget(QLabel("Prj Variant:"))
        sub_vbox3.addWidget(self.prj_variant)
        sub_vbox3.addWidget(QLabel("BOM list:"))
        sub_vbox3.addWidget(self.bom_list_search_path)
        sub_vbox3.addWidget(self.bom_list_search_select)

        self.q2_layout.addWidget(self.label_param)
        self.q2_layout.addLayout(top_hbox)
        self.q2_layout.addLayout(self.sub_top_q2_layout, 50)
        self.q2_layout.addLayout(sub_vbox3)
        self.q2_layout.addWidget(self.param_bom_list_view, 20)
        self.q2_layout.addWidget(self.log_panel, 30)

        l_logo = QLabel(LOGO)
        l_logo.setFont(QFont("MesloLGS NF", 10))
        l_logo.setAlignment(Qt.AlignCenter)
        l_version = QLabel("~ asterix24 ~\n- %s -\n" % MERGEBOM_VER)
        l_version.setFont(QFont("MesloLGS NF", 14))
        l_version.setAlignment(Qt.AlignCenter)

        self.main_layout.addWidget(l_logo)
        self.main_layout.addWidget(l_version)
        self.main_layout.addLayout(self.sub_layout)
        self.setLayout(self.main_layout)

        self.setWindowTitle("Merge BOM GUI")
        self.setGeometry(50, 50, 640, 480)
        QApplication.setStyle(QStyleFactory.create('Fusion'))
        QApplication.setPalette(QApplication.style().standardPalette())

        # Init MergeBOM class
        self.logger = Report(self.log_panel)
        self.logger.info("Mergbom start..\n")
        self.config = CfgMergeBom(logger=self.logger)

    def __fill_variant(self, prj):
        idx = self.prj_variant.currentIndex()
        self.prj_variant.clear()
        self.prj_variant.addItem(NO_VARIANT)
        self.prj_variant.addItems(self.prj_and_data[prj][2])
        if idx > 0:
            self.prj_variant.setCurrentIndex(idx)

    @pyqtSlot()
    def __select_prj_variant(self):
        prjs = self.param_prj_list_view.selectedItems()
        if len(prjs) == 1:
            self.__on_click_list(prjs[0])

    @pyqtSlot()
    def __merge_bom_select(self):
        bom_list = [i.text() for i in self.param_bom_list_view.selectedItems()]

        self.__merge_bom_hook(bom_list)

    @pyqtSlot()
    def __merge_bom_all(self):
        bom_list = []
        for i in range(self.param_bom_list_view.count()):
            bom_list.append(self.param_bom_list_view.item(i).text())

        self.__merge_bom_hook(bom_list)

    def __merge_bom_hook(self, bom_list):
        # Get prj info from view table
        param = {}
        for i in range(self.param_table_view.rowCount()):
            param[self.param_table_view.item(i, 0).text(
            )] = self.param_table_view.item(i, 1).text()

        try:
            m = MergeBom(bom_list, self.config, is_csv=self.merge_only_csv.isChecked(),
                         logger=self.logger)
        except FileNotFoundError as file_excp:
            self.logger.error("Error while merging..\n")
            self.logger.error(file_excp)

        for bom in bom_list:
            if self.delete_merged.isChecked():
                for i in bom_list:
                    os.remove(i)

            out_dir_path = self.merge_same_dir_path.text()
            if self.merge_same_dir.isChecked():
                out_dir_path = os.path.dirname(bom)

            outfilename = self.merge_bom_outname.text()
            if outfilename == "" or outfilename is None:
                self.log_panel.appendPlainText("Invalid out file name.")
                return

            try:
                out = os.path.join(out_dir_path, outfilename)
                write_xls(m.merge(),
                          list(map(os.path.basename, bom_list)),
                          self.config,
                          out,
                          param,
                          diff=False,
                          statistics=m.statistics(),
                          headers=m.header_data())
            except Exception as merge_excp:
                self.logger.error("Error while merging..\n")
                self.logger.error(str(merge_excp))

    @pyqtSlot()
    def __deploy_prj_select(self):
        search_path = os.path.dirname(self.selected_file.text())
        wk_name = os.path.basename(self.selected_file.text())
        wk_name, _ = os.path.splitext(wk_name)
        for i in self.param_prj_list_view.selectedItems():
            self.__deploy_project(wk_name, i.text(), search_path)

    @pyqtSlot()
    def __deploy_all_prj(self):
        search_path = os.path.dirname(self.selected_file.text())
        wk_name = os.path.basename(self.selected_file.text())
        wk_name, _ = os.path.splitext(wk_name)
        for i in range(self.param_prj_list_view.count()):
            self.__deploy_project(wk_name,
                                  self.param_prj_list_view.item(i).text(), search_path)

    def __deploy_project(self, wk_name, project_name, search_path):
        """
        Deploy folder structure

        Customer
        |  - version.txt
        |  - Readme.txt
        | - MainPrj1
                |  - Hw-Rev.0
                    | - Progetto-rev.12
                        | - Gerber
                                | - pcb-progetto-rev.A
                                | - ..
                        | - schematics.pdf
                        | - ...
                        | - version.txt
                    | - Progetto3-rev.0
                        | - Gerber
                                | - pcb-progetto-rev.A
                                | - ..
                        | - schematics.pdf
                        | - ...
                        | - version.txt
        | - MainPrj2
                |  - Hw-Rev.1-Prototype
                    | - Progetto-rev.3
                        | - Gerber
                                | - pcb-progetto-rev.A
                                | - ..
                        | - schematics.pdf
                        | - ...
                        | - version.txt
                    | - Progetto3-rev.7
                        | - Gerber
                                | - pcb-progetto-rev.A
                                | - ..
                        | - schematics.pdf
                        | - ...
                        | - version.txt
        """
        customer_name = self.deploy_customer_name.text()
        if customer_name is None or customer_name == "":
            self.logger.error("Invalid customer name\n")
            return

        deploy_path = self.deploy_path.text()
        if deploy_path is None or deploy_path == "":
            self.logger.error("Invalid deploy path\n")
            return

        # Get prj info from view table
        param = {}
        for i in range(self.param_table_view.rowCount()):
            param[self.param_table_view.item(i, 0).text(
            )] = self.param_table_view.item(i, 1).text()

        # make directory tree
        status = param.get("prj_status", "")
        if status:
            status = "-%s" % status

        if self.deploy_rev_name.text() is None or self.deploy_rev_name.text() == "":
            self.logger.error("Invalid hw revision for deploing.\n")
            return

        name_hw = TEMPLATE_HW_DIR % (self.deploy_rev_name.text(), status)
        name_prj = TEMPLATE_PRJ_NAME % (param.get("prj_prefix", ""),
                                        project_name,
                                        param.get("prj_hw_ver", "NONE"))
        name_gerber = TEMPLATE_PCB_NAME % (param.get("prj_prefix", ""),
                                           project_name,
                                           param.get("prj_pcb", "NONE"))

        name_gerber = TEMPLATE_PCB_NAME % (param.get("prj_prefix", ""),
                                           project_name,
                                           param.get("prj_pcb", "NONE"))

        gerber_dir = os.path.join(deploy_path, customer_name, wk_name,
                                  name_hw, name_prj, "Gerber")

        prj_dir = os.path.join(deploy_path, customer_name, wk_name,
                               name_hw, name_prj)

        try:
            os.makedirs(gerber_dir)
        except FileExistsError:
            self.logger.warning("Deploy directory exists\n")

        for item in DEFAULT_PRJ_DIR:
            folder, dst, src = item
            dst_name = dst % (project_name, param.get("prj_hw_ver", "NONE"))
            src_name = os.path.join(folder, project_name, src)
            if folder != "Pdf":
                src_name = os.path.join(
                    folder, project_name, src % project_name)

            src_file = os.path.join(search_path, src_name)
            dst_file = os.path.join(prj_dir, dst_name)
            try:
                self.logger.info("Copy project files:\n")
                self.logger.info("From: %s\n" % src_file)
                self.logger.info("To: %s\n" % dst_file)
                shutil.copy2(src_file, dst_file)
            except FileNotFoundError as cp_excp:
                self.logger.error("Unable to copy file:%s\n" % cp_excp)

        for item in [(project_name, name_gerber),
                     ("%s_pannel" % project_name, "%s_pannel" % name_gerber)]:
            self.logger.info("Copy Gerber files:\n")
            self.logger.info("From: %s\n" % item[0])
            self.logger.info("To: %s\n" % dst_file)
            src_path = os.path.join(search_path, "Gerber", item[0])
            try:
                copyGerberZip(item[1], src_path, gerber_dir)
            except FileNotFoundError as cp_excp:
                self.logger.error("Unable find gerber file:%s\n" % cp_excp)

        # if present deploy variants
        for index in range(self.prj_variant.count()):
            variant = self.prj_variant.itemText(index)
            if variant == NO_VARIANT:
                continue

            for item in DEFAULT_PRJ_DIR_VARIANT:
                folder, dst, src = item
                dst_name = dst % (project_name, variant,
                                  param.get("prj_hw_ver", "NONE"))
                src_name = os.path.join(folder, variant, src)
                if folder != "Pdf":
                    if "assembly" in dst:
                        src_name = os.path.join(
                            folder, variant, src % project_name)
                    else:
                        src_name = os.path.join(
                            folder, variant, src % (project_name, variant))

                dst_dir = os.path.join(prj_dir, variant)
                try:
                    os.makedirs(dst_dir)
                except FileExistsError:
                    self.logger.warning("Deploy directory exists\n")

                src_file = os.path.join(search_path, src_name)
                dst_file = os.path.join(dst_dir, dst_name)
                try:
                    self.logger.info("Copy project files:\n")
                    self.logger.info("From: %s\n" % src_file)
                    self.logger.info("To: %s\n" % dst_file)
                    shutil.copy2(src_file, dst_file)
                except FileNotFoundError as cp_excp:
                    self.logger.error("Unable to copy file:%s\n" % cp_excp)

    @pyqtSlot()
    def __autoname_out_file(self):
        if self.merge_autoname.isChecked():
            self.merge_bom_outname.setEnabled(False)
        else:
            self.merge_bom_outname.setText("merged_bom.xlsx")
            self.merge_bom_outname.setEnabled(True)

    @pyqtSlot()
    def __filter_checkbox(self):
        prjs = self.param_prj_list_view.selectedItems()
        if len(prjs) == 1:
            # fill variant combobox
            self.__on_click_list(prjs[0])
        else:
            if self.merge_only_csv.isChecked():
                csv_l = []
                for i in self.tmp_bom_list:
                    _, ext = os.path.splitext(i)
                    if ext == ".csv":
                        csv_l.append(i)
                self.param_bom_list_view.clear()
                self.param_bom_list_view.addItems(csv_l)
            else:
                self.param_bom_list_view.addItems(self.tmp_bom_list)

    @pyqtSlot()
    def __filter_prj(self):
        self.param_bom_list_view.clear()
        prjs = self.param_prj_list_view.selectedItems()
        if len(prjs) == 1:
            # fill variant combobox
            self.__on_click_list(prjs[0])

    @pyqtSlot()
    def __check_same_dir(self):
        if self.merge_same_dir.isChecked():
            self.merge_same_dir_path.setEnabled(False)
            self.merge_same_dir_path_select.setEnabled(False)
        else:
            self.merge_same_dir_path.setEnabled(True)
            self.merge_same_dir_path_select.setEnabled(True)

    @pyqtSlot()
    def __select_bom_search_path(self):
        root = os.path.dirname(self.selected_file.text())
        if self.bom_list_search_path is not None or self.bom_list_search_path.text() != "":
            root = self.bom_list_search_path.text()
        app = FileDialog(mode="dir", rootpath=root)
        d = app.directory()
        if d is not None and d != "":
            self.bom_list_search_path.setText(d)
            self.__on_click_list(self.param_prj_list_view.selectedItems())

    @pyqtSlot()
    def __select_merge_path(self):
        app = FileDialog(mode="dir", rootpath=self.merge_same_dir_path.text())
        d = app.directory()
        if d is not None and d != "":
            self.merge_same_dir_path.setText(d)

    @pyqtSlot()
    def __select_deploy_path(self):
        app = FileDialog(mode="dir", rootpath=self.deploy_path.text())
        d = app.directory()
        if d is not None and d != "":
            self.deploy_path.setText(d)

    @pyqtSlot()
    def __select_wk_file(self):
        app = FileDialog(mode="multiopen", rootpath=self.selected_file.text())
        line = app.selection()
        file_type = app.file_type()

        if line is not None and len(line) > 0:
            if file_type == SUPPORTED_FILE_TYPE[0]:  # workspace
                self.param_prj_list_view.clear()
                if len(line) > 1:
                    self.logger.warning(
                        "You should select only one Altium Workspace\n")
                    return

                line = line[0]
                self.selected_file.setText(line)
                self.__update_src_panel(line)
                self.merge_cmd_box.setEnabled(True)
                self.deploy_cmd_box.setEnabled(True)

            elif file_type == SUPPORTED_FILE_TYPE[1]:  # altium prj
                root_path = os.path.dirname(line[0])
                self.selected_file.setText(root_path)
                for prj in line:
                    root, _ = os.path.splitext(os.path.basename(prj))
                    n, d = get_parameterFromPrj(root, prj)
                    l = get_variantFromPrj(root, prj)
                    if n != "":
                        self.prj_and_data[n] = [root_path, d, l]

                        new_item = True
                        for i in range(self.param_prj_list_view.count()):
                            if self.param_prj_list_view.item(i).text() == n:
                                new_item = False

                        if new_item:
                            self.param_prj_list_view.addItem(n)

                self.merge_cmd_box.setEnabled(True)
                self.deploy_cmd_box.setEnabled(True)
                self.label_param.setText(
                    "Altium Project: %s" % os.path.basename(line[0].upper()))

            elif file_type == SUPPORTED_FILE_TYPE[2]:  # bom files
                self.tmp_bom_list = []
                self.param_prj_list_view.clear()
                root_path = os.path.dirname(line[0])
                self.selected_file.setText(root_path)

                self.param_table_view.clearContents()
                self.param_table_view.setRowCount(len(DEFAULT_PRJ_PARAM_DICT))
                for n, i in enumerate(DEFAULT_PRJ_PARAM_DICT):
                    tt = QTableWidgetItem(i)
                    tt.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    self.param_table_view.setItem(n, 0, tt)
                    self.param_table_view.setItem(n, 1,
                                                  QTableWidgetItem(DEFAULT_PRJ_PARAM_DICT[i]))

                self.tmp_bom_list = line
                self.merge_cmd_box.setEnabled(True)
                self.merge_only_csv.setChecked(False)
                self.label_param.setText(
                    "BOM Files path: %s" % root_path.upper())

            else:
                self.logger.warning("Unsupport file type.")

    def __update_src_panel(self, line):
        if line is None:
            self.logger.info("Workspace path is invalid or None\n")
            return

        self.param_prj_list_view.clear()
        root, _ = os.path.splitext(os.path.basename(line))
        self.label_param.setText("Workspace: %s" % root.upper())

        data = extrac_projects(line)
        root_path = os.path.dirname(line)
        for prj in data:
            n, d = get_parameterFromPrj(
                prj[0], os.path.join(root_path, prj[1]))
            l = get_variantFromPrj(prj[0], os.path.join(root_path, prj[1]))
            if n != "":
                self.prj_and_data[n] = [root_path, d, l]
                self.param_prj_list_view.addItem(n)

    def __on_click_list(self, item):
        current_prj = None
        if type(item) == str:
            current_prj = item
        elif type(item) == list:
            if len(item) == 0:
                return
            current_prj = item[0].text()
        else:
            current_prj = item.text()

        if current_prj == "":
            return

        d = self.prj_and_data[current_prj][1]
        self.param_table_view.clearContents()
        self.param_table_view.setRowCount(len(d))
        for n, i in enumerate(d):
            tt = QTableWidgetItem(i)
            tt.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.param_table_view.setItem(n, 0, tt)
            self.param_table_view.setItem(n, 1, QTableWidgetItem(d[i]))

        self.param_table_view.resizeRowsToContents()
        self.param_bom_list_view.clear()
        root = self.prj_and_data[current_prj][0]

        flt = "*.xlsx"
        if self.merge_only_csv.isChecked():
            flt = "*.csv"

        # Get prj info from view table
        param = {}
        for i in range(self.param_table_view.rowCount()):
            param[self.param_table_view.item(i, 0).text(
            )] = self.param_table_view.item(i, 1).text()

        variant = self.prj_variant.currentText()
        search_dirs = [root, os.path.join(
            root, "Assembly"), os.path.join(root, "..", "Assembly")]
        if self.merge_prj_files.isChecked():
            search_dirs = [os.path.join(root, "Assembly", current_prj), os.path.join(
                root, "..", "Assembly", current_prj)]
            if variant != NO_VARIANT:
                search_dirs = [os.path.join(root, "Assembly", current_prj, variant),
                               os.path.join(root, "Assembly", variant),
                               os.path.join(root, "..", "Assembly",
                                            current_prj, variant),
                               os.path.join(root, "Assembly", "%s_var-%s" % (current_prj, variant))]

        if self.bom_list_search_path is None or self.bom_list_search_path.text() != "":
            search_dirs = [self.bom_list_search_path.text()]

        d_boms = {}
        for item in search_dirs:
            self.logger.info("Search bom files in: %s\n" % item)
            for path in Path(item).rglob(flt):
                d_boms[os.path.join(path.parent, path.name)] = ""

        if self.merge_autoname.isChecked():
            self.merge_bom_outname.setText(MERGED_FILE_TEMPLATE % (
                current_prj, param.get("prj_hw_ver", "NOHWVER")))
            if variant != NO_VARIANT:
                self.merge_bom_outname.setText(
                    MERGED_FILE_TEMPLATE_VARIANT % (current_prj, variant, param.get("prj_hw_ver", "NOHWVER")))

        self.param_bom_list_view.addItems(d_boms.keys())
        self.__fill_variant(current_prj)


FILE_FILTERS = ";;".join(SUPPORTED_FILE_TYPE)


class FileDialog(QWidget):
    """
    Common Dialog to select file from filesystem
    """

    def __init__(self, title="Open File", mode='open', rootpath=os.path.expanduser("~")):
        super().__init__()
        self.title = title
        self.mode = mode
        self.rootpath = rootpath

        self.file_path = None
        self.filetype = None
        self.dir_name = None

        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480

        self.init_ui()

    def init_ui(self):
        """
        init
        """
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        if self.mode == 'multiopen':
            self.__open_file_names_dialog()
        elif self.mode == 'dir':
            self.__open_directory_dialog()
        else:
            self.__open_file_name_dialog()

        self.show()

    def selection(self):
        """
        Return selected path
        """
        return self.file_path

    def file_type(self):
        """
        Return selected path
        """
        return self.filetype

    def directory(self):
        """
        Return selected path
        """
        return self.dir_name

    def __open_directory_dialog(self):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        dirname = QFileDialog.getExistingDirectory(
            self, self.title, self.rootpath, options=options)
        if dirname:
            self.dir_name = dirname

    def __open_file_name_dialog(self):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        filename, filter_type = QFileDialog.getOpenFileName(
            self, self.title, self.rootpath, FILE_FILTERS, options=options)

        if filename:
            self.file_path = filename
            self.filetype = filter_type

    def __open_file_names_dialog(self):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        files, filter_type = QFileDialog.getOpenFileNames(
            self, self.title, self.rootpath, FILE_FILTERS, options=options)
        if files:
            self.file_path = files
            self.filetype = filter_type


def main_app():
    """
    MergeBom Launcher
    """

    app = QApplication(sys.argv)
    gallery = MergeBomGUI()
    gallery.show()
    sys.exit(app.exec_())
