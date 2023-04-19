from PyQt6.QtWidgets import QDialog
from gui.dlgQueryTool import Ui_Dialog

class QueryToolDialog(QDialog, Ui_Dialog):
    def __init__(self):
        super().__init__()

        self.setupUi(self)
