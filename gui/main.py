from PyQt5.QtCore import QDateTime, Qt, QTimer
from PyQt5.QtWidgets import (QApplication, QCheckBox, QComboBox, QDateTimeEdit,
                             QDial, QDialog, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                             QProgressBar, QPushButton, QRadioButton, QScrollBar, QSizePolicy,
                             QSlider, QSpinBox, QStyleFactory, QTableWidget, QTabWidget, QTextEdit,
                             QVBoxLayout, QWidget, QInputDialog, QLineEdit, QFileDialog)

from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot

import sys


class MergeBomGUI(QDialog):
    """
    GUI frontend of mergebom script.
    With the GUI is possible also to deploy the proct on given directory.
    """

    def __init__(self, parent=None):
        super(MergeBomGUI, self).__init__(parent)

        styleComboBox = QComboBox()
        styleComboBox.addItems(QStyleFactory.keys())

        styleLabel = QLabel("&Style:")
        styleLabel.setBuddy(styleComboBox)

        defaultPushButton = QPushButton("Default Push Button")
        defaultPushButton.setDefault(True)
        defaultPushButton.clicked.connect(self.__on_click)

        topLayout = QHBoxLayout()
        topLayout.addWidget(styleLabel)
        topLayout.addWidget(styleComboBox)
        topLayout.addWidget(defaultPushButton)

        mainLayout = QGridLayout()
        mainLayout.addLayout(topLayout, 0, 0, 1, 2)
        self.setLayout(mainLayout)

        self.setWindowTitle("Merge BOM GUI")
        QApplication.setStyle(QStyleFactory.create('Fusion'))
        QApplication.setPalette(QApplication.style().standardPalette())

    @pyqtSlot()
    def __on_click(self):
        app = FileDialog()
        print(app.selection)


FILE_FILTERS = {
    'workspace': "Altium WorkSpace (*.DsnWrk);;Altium Project (*.PrjPcb);;All Files (*)"
}


class FileDialog(QWidget):
    """
    Common Dialog to select file from filesystem
    """

    def __init__(self, title="Open File", ext_filter="workspace", mode='open'):
        super().__init__()
        self.title = title
        self.mode = mode
        self.filter = ext_filter

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
            self, self.title, "", FILE_FILTERS[self.filter], options=options)
        if filename:
            self.file_path = filename

    def __open_file_names_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(
            self, self.title, "", FILE_FILTERS[self.filter], options=options)
        if files:
            self.file_path = files

    def __save_file_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getSaveFileName(
            self, self.title, "", FILE_FILTERS[self.filter], options=options)
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
