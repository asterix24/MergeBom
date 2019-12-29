#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
from lib.cfg import LOGO, extrac_projects, get_parameterFromPrj, find_bomfiles

from PyQt5.QtCore import QDateTime, Qt, pyqtSlot
from PyQt5.QtWidgets import (QApplication, QPlainTextEdit, QComboBox, QDateTimeEdit, QDialog, QGroupBox, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QRadioButton, QScrollBar, QSlider, QSpinBox,
                             QStyleFactory, QTableWidget, QVBoxLayout, QWidget, QInputDialog, QLineEdit,
                             QFileDialog, QListWidget, QTableWidgetItem, QHeaderView)

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
        self.sub_layout.addLayout(self.q1_layout)
        self.sub_layout.addLayout(self.q2_layout)

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
        self.q2_layout.addWidget(self.param_prj_list_view, 25)
        self.q2_layout.addWidget(self.param_table_view, 75)


        l_logo = QLabel(LOGO)
        l_logo.setFont(QFont("MesloLGS NF", 10))
        l_logo.setAlignment(Qt.AlignCenter)

        self.log_panel = QPlainTextEdit(self)
        self.log_panel.setEnabled(False)
        self.log_layout.addWidget(self.log_panel)

        self.main_layout.addWidget(l_logo)
        self.main_layout.addLayout(self.sub_layout)
        self.main_layout.addLayout(self.log_layout)
        self.setLayout(self.main_layout)

        self.setWindowTitle("Merge BOM GUI")
        QApplication.setStyle(QStyleFactory.create('Fusion'))
        QApplication.setPalette(QApplication.style().standardPalette())

    @pyqtSlot()
    def __merge_bom_select(self):
        self.log_panel.setEnabled(True)
        pass

    @pyqtSlot()
    def __merge_bom_all(self):
        pass

    @pyqtSlot()
    def __deploy_bom_select(self):
        pass

    @pyqtSlot()
    def __deploy_bom_all(self):
        pass

    @pyqtSlot()
    def __select_wk_file(self):
        app = FileDialog(rootpath=self.wk_path.text())
        line = app.selection()
        print(line)
        wk_file_name, _ = os.path.splitext(os.path.basename(line))
        if line is not None:
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
        l = find_bomfiles(root, current_prj, True)
        self.param_bom_list_view.addItems(l)
        l = find_bomfiles(os.path.join(root, "Assembly"), current_prj, False)
        self.param_bom_list_view.addItems(l)



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
