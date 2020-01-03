#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
from lib.cfg import LOGO, extrac_projects, get_parameterFromPrj, find_bomfiles
from lib.cfg import MERGED_FILE_TEMPLATE
from lib.cfg import CfgMergeBom
from lib.report import Report, write_xls
from mergebom_class import *

from PyQt5.QtCore import QDateTime, Qt, pyqtSlot
from PyQt5.QtWidgets import (QApplication, QPlainTextEdit, QCheckBox, QDialog, QGroupBox,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, QStyleFactory,
                             QTableWidget, QVBoxLayout, QWidget, QFileDialog, QListWidget,
                             QTableWidgetItem, QHeaderView)

from PyQt5.QtGui import QIcon, QFont


class MergeBomGUI(QDialog):
    """
    GUI frontend of mergebom script.
    With the GUI is possible also to deploy the proct on given directory.
    """

    def __init__(self, parent=None):
        super(MergeBomGUI, self).__init__(parent)

        self.prj_and_data = {}
        self.param_prj_list_view = None
        self.param_bom_list_view = None


        self.main_layout = QVBoxLayout()
        self.sub_layout = QHBoxLayout()
        self.q1_layout = QVBoxLayout()
        self.q2_layout = QVBoxLayout()
        self.log_layout = QVBoxLayout()
        self.sub_layout.addLayout(self.q1_layout, 15)
        self.sub_layout.addLayout(self.q2_layout, 85)

        # Q1 right quadrant
        self.wk_path = QLineEdit(os.path.expanduser('~/'))
        btn = QPushButton("Select")
        btn.setDefault(True)

        self.merge_sel = QPushButton("Merge Select")
        self.merge_sel.setDefault(True)
        self.merge_sel.clicked.connect(self.__merge_bom_select)
        self.merge_all = QPushButton("Merge All")
        self.merge_all.setDefault(True)
        self.merge_all.clicked.connect(self.__merge_bom_all)

        self.merge_only_csv = QCheckBox("CSV Filter", self)
        self.merge_only_csv.setChecked(True)
        self.merge_only_csv.stateChanged.connect(self.__filterCheckbox)
        self.merge_same_dir = QCheckBox("Merge in same directory", self)
        self.merge_same_dir.stateChanged.connect(self.__check_same_dir)
        self.merge_same_dir_path = QLineEdit(os.path.expanduser('~/'))
        self.delete_merged = QCheckBox("Delete Mergeded", self)
        self.merge_bom_outname_label = QLabel("Merged out file name:")
        self.merge_bom_outname = QLineEdit("merged_bom.xlsx")
        self.merge_bom_outname.setEnabled(False)
        self.merge_autoname = QCheckBox("Autoname out file", self)
        self.merge_autoname.setChecked(True)
        self.merge_autoname.stateChanged.connect(self.__autoname_out_file)

        self.deploy_sel = QPushButton("Deploy Select")
        self.deploy_sel.setDefault(True)
        self.deploy_sel.clicked.connect(self.__deploy_bom_select)
        self.deploy_all = QPushButton("Deploy All")
        self.deploy_all.setDefault(True)
        self.deploy_all.clicked.connect(self.__deploy_bom_all)

        self.merge_cmd_box = QGroupBox("MergeBOM Commands")
        self.merge_cmd_box.setEnabled(False)
        vbox = QVBoxLayout()
        self.merge_cmd_box.setLayout(vbox)
        vbox.addWidget(self.merge_sel)
        vbox.addWidget(self.merge_all)
        vbox.addWidget(self.merge_only_csv)
        vbox.addWidget(self.merge_same_dir)
        vbox.addWidget(self.merge_same_dir_path)
        vbox.addWidget(self.delete_merged)
        vbox.addWidget(self.merge_bom_outname_label)
        vbox.addWidget(self.merge_bom_outname)
        vbox.addWidget(self.merge_autoname)

        self.deploy_cmd_box = QGroupBox("Deploy Commands")
        self.deploy_cmd_box.setEnabled(False)
        vbox = QVBoxLayout()
        self.deploy_cmd_box.setLayout(vbox)
        vbox.addWidget(self.deploy_sel)
        vbox.addWidget(self.deploy_all)

        btn.clicked.connect(self.__select_wk_file)
        self.q1_layout.addWidget(self.wk_path)
        self.q1_layout.addWidget(btn)
        self.q1_layout.addWidget(self.merge_cmd_box)
        self.q1_layout.addWidget(self.deploy_cmd_box)

        self.q1_layout.addStretch(1)

        # Q2 left quadrant
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

        self.label_param = QLabel("Workspace: --")
        self.q2_layout.addWidget(self.label_param)
        self.q2_layout.addWidget(QLabel("Project list:"))
        self.q2_layout.addWidget(self.param_prj_list_view, 25)
        self.q2_layout.addWidget(QLabel("Project params:"))
        self.q2_layout.addWidget(self.param_table_view, 60)
        self.q2_layout.addWidget(QLabel("BOM list:"))
        self.q2_layout.addWidget(self.param_bom_list_view, 15)


        l_logo = QLabel(LOGO)
        l_logo.setFont(QFont("MesloLGS NF", 10))
        l_logo.setAlignment(Qt.AlignCenter)

        self.log_panel = QPlainTextEdit(self)
        self.log_panel.setReadOnly(True)
        self.log_layout.addWidget(self.log_panel)

        self.main_layout.addWidget(l_logo)
        self.main_layout.addLayout(self.sub_layout)
        self.main_layout.addLayout(self.log_layout)
        self.setLayout(self.main_layout)

        self.setWindowTitle("Merge BOM GUI")
        self.setGeometry(50, 50, 640, 480)
        QApplication.setStyle(QStyleFactory.create('Fusion'))
        QApplication.setPalette(QApplication.style().standardPalette())

        # Init MergeBOM class
        self.logger = Report(terminal=False)
        #logger.write_logo()

        self.config = CfgMergeBom(logger=self.logger)

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
            param[self.param_table_view.item(i, 0).text()] = self.param_table_view.item(i, 1).text()

        m = MergeBom(bom_list, self.config, is_csv=self.merge_only_csv.isChecked(),
                     logger=self.logger)

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

            out = os.path.join(out_dir_path, outfilename)
            write_xls(m.merge(),
                     list(map(os.path.basename, bom_list)),
                     self.config,
                     out,
                     param,
                     diff=False,
                     statistics=m.statistics())

    @pyqtSlot()
    def __deploy_bom_select(self):
        pass

    @pyqtSlot()
    def __deploy_bom_all(self):
        pass

    @pyqtSlot()
    def __autoname_out_file(self):
        if not self.merge_autoname.isChecked():
            self.merge_bom_outname.setText("merged_bom.xlsx")
            self.merge_bom_outname.setEnabled(True)
        else:
            self.merge_bom_outname.setEnabled(False)

    @pyqtSlot()
    def __filterCheckbox(self):
        prjs = self.param_prj_list_view.selectedItems()
        if len(prjs) == 1:
            self.__on_click_list(prjs[0])

    @pyqtSlot()
    def __check_same_dir(self):
        if self.merge_same_dir.isChecked():
            self.merge_same_dir_path.setEnabled(False)
        else:
            self.merge_same_dir_path.setEnabled(True)

    @pyqtSlot()
    def __select_wk_file(self):
        app = FileDialog(rootpath=self.wk_path.text())
        line = app.selection()

        if line is not None and line != "":
            wk_file_name, _ = os.path.splitext(os.path.basename(line))
            self.wk_path.setText(line)
            self.__update_src_panel(line)
            self.merge_cmd_box.setEnabled(True)
            self.deploy_cmd_box.setEnabled(True)

    def __update_src_panel(self, line):
        if line is None:
            print("Workspace path is invalid or None")
            return

        root, ext = os.path.splitext(os.path.basename(line))
        self.label_param.setText("Workspace: %s" % root)

        data = extrac_projects(line)
        root_path = os.path.dirname(line)
        for prj in data:
            n, d = get_parameterFromPrj(
                prj[0], os.path.join(root_path, prj[1]))
            self.prj_and_data[n] = [root_path, d]
            self.param_prj_list_view.addItem(n)

    def __on_click_list(self, item):
        current_prj = None
        if type(item) == str:
            current_prj = item
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

        self.param_bom_list_view.clear()
        root = self.prj_and_data[current_prj][0]

        for pths in [root, os.path.join(root, "Assembly")]:
            l = find_bomfiles(pths, current_prj,
                              self.merge_only_csv.isChecked())
            self.param_bom_list_view.addItems(l[1])

        self.merge_bom_outname.setText(MERGED_FILE_TEMPLATE % current_prj)



FILE_FILTERS = {
    'workspace': "Altium WorkSpace (*.DsnWrk);;Altium Project (*.PrjPCB);;All Files (*)"
}


class FileDialog(QWidget):
    """
    Common Dialog to select file from filesystem
    """

    def __init__(self, title="Open File", ext_filter="workspace", mode='open', rootpath=os.path.expanduser("~")):
        super().__init__()
        self.title = title
        self.mode = mode
        self.filter = ext_filter
        self.rootpath = rootpath

        self.file_path = None

        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        if self.mode == 'save':
            self.__save_file_dialog()
        elif self.mode == 'multiopen':
            self.__open_file_names_dialog()
        else:
            self.__open_file_name_dialog()

        self.show()

    def selection(self):
        """
        Return selected path
        """
        return self.file_path

    def __open_file_name_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getOpenFileName(
            self, self.title, self.rootpath, FILE_FILTERS[self.filter], options=options)
        if filename:
            self.file_path = filename

    def __open_file_names_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(
            self, self.title, self.rootpath, FILE_FILTERS[self.filter], options=options)
        if files:
            self.file_path = files

    def __save_file_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getSaveFileName(
            self, self.title, self.rootpath, FILE_FILTERS[self.filter], options=options)
        if filename:
            self.file_path = filename


def main_app():
    """
    MergeBom Launcher
    """

    app = QApplication(sys.argv)
    gallery = MergeBomGUI()
    gallery.show()
    sys.exit(app.exec_())