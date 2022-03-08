#!/usr/bin/env python3


import sys
import os
import traceback

from PyQt5.QtWidgets import *
from PyQt5.QtCore    import *
from PyQt5.QtGui     import *

import tid_ss_lib.navigate
import tid_ss_lib.project_sheet
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

        if 'SMARTSHEETS_API' in os.environ:
            api = os.environ['SMARTSHEETS_API']
        else:
            api = ''

        self.api_key = QLineEdit(api)
        fl.addRow('API Key',self.api_key)

        api_help = QTextEdit()
        api_help.setPlainText('To generate an API key see:\nhttps://help.smartsheet.com/articles/2482389-generate-API-key')
        api_help.setReadOnly(True)
        fl.addRow('',api_help)

        refreshRun = QPushButton("Refresh Project List")
        refreshRun.pressed.connect(self.refreshPressed)
        fl.addRow("",refreshRun)

        self.proj_list = QComboBox()
        self.projects = {0: ''}
        self.proj_list.addItem('Manual Entry')
        self.proj_list.currentIndexChanged.connect(self.proj_select)
        fl.addRow('Project Select',self.proj_list)

        self.folder_id = QLineEdit()
        fl.addRow('Folder ID',self.folder_id)

        id_help = QTextEdit()
        id_help.setPlainText("Folder ID to check/fix.\nRight click on project folder,\nselect properties and copy 'Folder ID'")
        id_help.setReadOnly(True)
        fl.addRow('',id_help)

        self.do_fixes = QCheckBox()
        fl.addRow('Do Fixes',self.do_fixes)

        fix_help = QTextEdit()
        fix_help.setPlainText("Check Do Fixes to apply fixes,\notherwise the script will just report problems.")
        fix_help.setReadOnly(True)
        fl.addRow('',fix_help)

        scriptRun = QPushButton("Execute Script")
        scriptRun.pressed.connect(self.fixPressed)
        fl.addRow("",scriptRun)

        closeBtn  = QPushButton("Close")
        closeBtn.pressed.connect(self.close)
        fl.addRow("",closeBtn)

        self.resize(500,500)

    @pyqtSlot(int)
    def proj_select(self, index):
        if index >= 0:
            self.folder_id.setText(str(self.projects[index]))

    @pyqtSlot()
    def fixPressed(self):
        try:
            print("")
            client = smartsheet.Smartsheet(self.api_key.text())
            folder = int(self.folder_id.text())
            doFixes = self.do_fixes.isChecked()
            tid_ss_lib.navigate.check_project(client=client,folderId=folder, doFixes=doFixes)
            print("Done!")
        except Exception as msg:
            traceback.print_exc()
            print(f"\n\n\nGot Error:\n{msg}\n\n")


    @pyqtSlot()
    def refreshPressed(self):
        try:
            client = smartsheet.Smartsheet(self.api_key.text())
            lst = tid_ss_lib.project_sheet.get_project_list(client=client)

            self.proj_list.clear()
            self.projects = {0: ''}
            self.proj_list.addItem('Manual Entry')

            for i,l in enumerate(lst):
                self.projects[i+1] = l['id']
                self.proj_list.addItem(l['name'] + ' - ' + l['pm'])

        except Exception as msg:
            print(f"\n\n\nGot Error:\n{msg}\n\n")

appTop = QApplication(sys.argv)

gui = ProjectFix()
gui.show()
appTop.exec_()

