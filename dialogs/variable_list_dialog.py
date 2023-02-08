import sys
from PyQt6.QtWidgets import QApplication, QDialog
from PyQt6.QtCore import Qt
from gui.frmVariableListDialog import Ui_Dialog

class VariableListDialog(QDialog, Ui_Dialog):
    def __init__(self):
        super().__init__()

        #load ui
        self.setupUi(self)

        self.init_variable_list()

    def init_variable_list(self):
        self.tbl_variables.setColumnCount(3)
        self.tbl_variables.setHorizontalHeaderLabels(["Variable Name", "Variable Label", ""])




