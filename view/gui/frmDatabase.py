# Form implementation generated from reading ui file 'view\ui\frmDatabase.ui'
#
# Created by: PyQt6 UI code generator 6.4.0
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_frmDatabase(object):
    def setupUi(self, frmDatabase):
        frmDatabase.setObjectName("frmDatabase")
        frmDatabase.resize(451, 349)
        self.verticalLayout = QtWidgets.QVBoxLayout(frmDatabase)
        self.verticalLayout.setContentsMargins(3, 3, 3, 3)
        self.verticalLayout.setObjectName("verticalLayout")
        self.table_database = QtWidgets.QTableView(frmDatabase)
        self.table_database.setObjectName("table_database")
        self.verticalLayout.addWidget(self.table_database)

        self.retranslateUi(frmDatabase)
        QtCore.QMetaObject.connectSlotsByName(frmDatabase)

    def retranslateUi(self, frmDatabase):
        _translate = QtCore.QCoreApplication.translate
        frmDatabase.setWindowTitle(_translate("frmDatabase", "Database"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    frmDatabase = QtWidgets.QWidget()
    ui = Ui_frmDatabase()
    ui.setupUi(frmDatabase)
    frmDatabase.show()
    sys.exit(app.exec())
