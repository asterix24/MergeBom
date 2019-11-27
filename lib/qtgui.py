#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
from lib.cfg import LOGO, extrac_projects, get_parameterFromPrj, find_bomfiles

from PyQt5.QtCore import QDateTime, Qt
from PyQt5.QtWidgets import (QApplication, QComboBox, QDateTimeEdit, QDialog, QGroupBox, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QRadioButton, QScrollBar, QSlider, QSpinBox,
                             QStyleFactory, QTableWidget, QVBoxLayout, QWidget, QInputDialog, QLineEdit,
                             QFileDialog, QListWidget, QTableWidgetItem, QHeaderView)

from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot


class MergeBomGUI(QDialog):
    """
    GUI frontend of mergebom script.
    With the GUI is possible also to deploy the proct on given directory.
    """

    def __init__(self, parent=None):
        super(MergeBomGUI, self).__init__(parent)

        self.prj_and_data = {}
        self.param_prj_list = None
        self.param_prj_list_view = None
        self.param_table_view = None

        self.source_data = None
        self.src_data = None
        self.src_data_name = None
        self.src_data_name_lbl = None

        self.altium_group = None
        self.wk_path = None

        self.param_data = None
        self.headers = ["Name", "Value"]
        self.rows = [('-', '-')]

        self.__altium_group()
        self.__param_data()
        self.__source_data()

        main_layout = QVBoxLayout()
        main_layout.addWidget(QLabel(LOGO))
        main_layout.addWidget(self.altium_group)
        main_layout.addWidget(self.source_data)
        main_layout.addWidget(self.param_data)
        self.setLayout(main_layout)

        self.setWindowTitle("Merge BOM GUI")
        QApplication.setStyle(QStyleFactory.create('Fusion'))
        QApplication.setPalette(QApplication.style().standardPalette())

    @pyqtSlot()
    def __on_click(self):
        app = FileDialog(rootpath=self.wk_path.text())
        line = app.selection()
        print(line)
        if line is not None:
            self.wk_path.setText(line)
            self.__update_src_panel()

    def __update_src_panel(self):
        line = self.wk_path.text()
        if line is None:
            print("Workspace path is invalid or None")
            return

        root, ext = os.path.splitext(
            os.path.basename(line))

        lbl_root = "Workspace Name:"
        data_type = "Altium Workspace"
        if ext in "PrjPCB":
            data_type = "Altium Project"
            lbl_root = "Project Name:"

        self.src_data.insertItem(0, data_type)
        self.src_data.setCurrentIndex(0)

        self.src_data_name.setText(root)
        self.src_data_name_lbl.setText(lbl_root)

        data = extrac_projects(line)
        root_path = os.path.dirname(line)
        for prj in data:
            n, d = get_parameterFromPrj(
                prj[0], os.path.join(root_path, prj[1]))
            print(d, n)
            self.prj_and_data[n] = d
            self.param_prj_list_view.addItem(n)

    def __update_prj_panel(self):
        pass

    def __source_data(self):
        self.source_data = QGroupBox("Source Data")

        self.src_data = QComboBox()
        self.src_data.addItems(["Custom"])
        srclbl = QLabel("&From:")
        srclbl.setBuddy(self.src_data)

        hsrc = QHBoxLayout()
        hsrc.addWidget(srclbl)
        hsrc.addWidget(self.src_data)

        self.src_data_name = QLabel("None")
        self.src_data_name_lbl = QLabel("Data Name:")

        info = QHBoxLayout()
        info.addWidget(self.src_data_name_lbl)
        info.addWidget(self.src_data_name)

        data = QVBoxLayout()
        data.addLayout(hsrc)
        data.addLayout(info)

        self.source_data.setLayout(data)

    def __on_click_list(self, item):
        current_prj = None
        if type(item) == str:
            current_prj = item
        else:
            current_prj = item.text()

        if current_prj == "":
            return

        d = self.prj_and_data[current_prj]
        self.param_table_view.clearContents()
        self.param_table_view.setRowCount(len(d))
        for n, i in enumerate(d):
            tt = QTableWidgetItem(i)
            tt.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.param_table_view.setItem(n, 0, tt)
            self.param_table_view.setItem(n, 1, QTableWidgetItem(d[i]))

    def __param_data(self):
        self.param_data = QGroupBox("Project and Settings")
        self.param_table_view = QTableWidget()
        hdr = self.param_table_view.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Stretch)
        self.param_table_view.setRowCount(1)
        self.param_table_view.setColumnCount(2)
        self.param_table_view.setHorizontalHeaderLabels(["Param", "Value"])

        self.param_prj_list_view = QListWidget()
        self.param_prj_list_view.itemClicked.connect(self.__on_click_list)
        self.param_prj_list_view.currentTextChanged.connect(
            self.__on_click_list)

        hbox = QHBoxLayout()
        hbox.addWidget(self.param_prj_list_view)
        hbox.addWidget(self.param_table_view)

        self.param_data.setLayout(hbox)

    def __altium_group(self):
        self.altium_group = QGroupBox("Select Altium Wks or Prj")

        self.wk_path = QLineEdit(os.path.expanduser('~/'))
        btn = QPushButton("Select")
        btn.setDefault(True)
        btn.clicked.connect(self.__on_click)

        lfile = QHBoxLayout()
        lfile.addWidget(self.wk_path)
        lfile.addWidget(btn)

        self.altium_group.setLayout(lfile)


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
