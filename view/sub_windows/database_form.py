from PyQt6.QtWidgets import QMdiSubWindow, QWidget, QTableView, QVBoxLayout
from PyQt6.QtCore import QCoreApplication
from gui.frmDatabase import Ui_frmDatabase

class DatabaseMdiSubWindow(QMdiSubWindow):
    def __init__(self):
        super().__init__()

        widget = DatabaseQWidget()
        self.setWidget(widget)

class DatabaseQWidget(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(3, 3, 3, 3)
        layout.setObjectName("verticalLayout")
        table = QTableView()
        table.setObjectName("table_database")
        layout.addWidget(table)

        self.retranslateUi()

    def retranslateUi(self):
        _translate = QCoreApplication.translate
        self.setWindowTitle(_translate("frmDatabase", "Database"))