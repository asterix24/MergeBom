#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
from lib.cfg import LOGO

from PyQt5.QtCore import QDateTime, Qt, QAbstractTableModel, QVariant
from PyQt5.QtWidgets import (QApplication, QCheckBox, QComboBox, QDateTimeEdit, QTableView,
                             QDial, QDialog, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                             QProgressBar, QPushButton, QRadioButton, QScrollBar, QSizePolicy,
                             QSlider, QSpinBox, QStyleFactory, QTableWidget, QTabWidget, QTextEdit,
                             QVBoxLayout, QWidget, QInputDialog, QLineEdit, QFileDialog)

from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot


class MergeBomGUI(QDialog):
    """
    GUI frontend of mergebom script.
    With the GUI is possible also to deploy the proct on given directory.
    """

    def __init__(self, parent=None):
        super(MergeBomGUI, self).__init__(parent)

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

    def __param_data(self):
        self.param_data = QGroupBox("Parms")
        model = ParamDataModel(self.headers, self.rows)
        view = QTableView()
        view.setModel(model)

        hbox = QHBoxLayout()
        hbox.addWidget(view)

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


class ParamDataModel(QAbstractTableModel):
    """
    Model to display all mergebom param get from variuos source,
    like altium wk/prj or use insertion.
    """

    def __init__(self, headers=["Name", "Value"], rows=[("-", "-")]):
        super().__init__()
        self.rows = rows
        self.headers = headers

    def rowCount(self, parent):
        return len(self.rows)

    def columnCount(self, parent):
        return len(self.headers)

    def data(self, index, role):
        if role != Qt.DisplayRole:
            return QVariant()
        return self.rows[index.row()][index.column()]

    def headerData(self, section, orientation, role):
        if role != Qt.DisplayRole or orientation != Qt.Horizontal:
            return QVariant()
        return self.headers[section]


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
