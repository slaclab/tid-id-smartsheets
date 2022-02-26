#!/usr/bin/env python3

import sys

from PyQt5.QtWidgets import *
from PyQt5.QtCore    import *
from PyQt5.QtGui     import *

import tid_ss_lib.navigate
import smartsheet

class ProjectFix(QWidget):

    def __init__(self, parent=None):
        super(ProjectFix, self).__init__(parent)

        self.setWindowTitle("TID ID Smartsheets Project Fix")

        # Setup status widgets
        top = QVBoxLayout()
        self.setLayout(top)

        fl = QFormLayout()
        fl.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        fl.setRowWrapPolicy(QFormLayout.DontWrapRows)
        fl.setFormAlignment(Qt.AlignHCenter | Qt.AlignTop)
        fl.setLabelAlignment(Qt.AlignRight)
        top.addLayout(fl)

        self.api_key = QLineEdit()
        fl.addRow('API Key',self.api_key)

        self.folder_id = QLineEdit()
        fl.addRow('Folder ID',self.api_key)

        self.do_fixes = QCheckBox()
        fl.addRow('Do Fixes',self.do_fixes)

        scriptRun = QPushButton("Execute Script")
        scriptRun.pressed.connect(self.fixPressed)
        fl.addRow("",scriptRun)

        self.resize(500,600)

    @pyqtSlot()
    def fixPressed(self):
        client = smartsheet.Smartsheet(self.api_key)
        folder = int(self.folder_id)
        doFix = self.do_fixes.isChecked()
        tid_ss_lib.navigate.check_project(client=client,folderId=folder, doFixes=doFixes)

appTop = QApplication(sys.argv)

gui = gui.ProjectFix()
gui.show()
appTop.exec_()

