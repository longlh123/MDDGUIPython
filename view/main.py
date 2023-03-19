import os, sys
sys.path.append(os.getcwd())

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QHeaderView, QPushButton, QTreeWidgetItem, QAbstractItemView
from PyQt6.QtCore import Qt, QMimeData, QPoint
from PyQt6.QtGui import QIcon, QDrag, QColor
from gui.frmMain import Ui_MainWindow

from dialogs.variable_list_dialog import VariableListDialog
from objects.IOMObject import Questions, Question, QuestionTreeItem

from pathlib import Path
import json
import pickle
import win32com.client as w32
from objects.enumerations import dataTypeConstants, objectTypeConstants, objectDepartments

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, department):
        super().__init__()

        #Load the Ui
        self.setupUi(self) 

        self.mdd_path = ""
        self.department = department
        #self.questions = Questions()

        self.MDM = w32.Dispatch(r'MDM.Document')

        self.actionOpen.triggered.connect(self.open_file_dialog)
        
    def open_file_dialog(self):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setNameFilter("Data Collection Data Files (*.mdd)")
        dialog.setViewMode(QFileDialog.ViewMode.List)

        if dialog.exec():
            filenames = dialog.selectedFiles()

            if filenames:
                self.mdd_path = Path(filenames[0])

                QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
                
                self.init_questions()
                #self.init_bvc_questions()
                
                QApplication.restoreOverrideCursor()

    def init_questions(self):
        self.MDM.Open(str(self.mdd_path))

        self.tree_questions.clear()
        self.tree_questions.setHeaderLabel(self.mdd_path.name)
        
        #Set the drap and drop in QTreeWidget and QTableWidget
        self.tree_questions.setDragEnabled(True)
        self.tree_questions.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
        
        for field in self.MDM.Fields:
            if field.Name == "Phase_Campaigns":
                a = ""
            node = self.create_a_node(field)
            
            if node is not None:
                self.tree_questions.addTopLevelItem(node)
        
        self.MDM.Close()

        #Connect the 'itemPressed' signal to the 'handleStartDrap' method
        self.tree_questions.itemPressed.connect(self.hanldeItemPressed)
        
    def create_a_node(self, field, parent=None):
        if str(field.ObjectTypeValue) == objectTypeConstants.mtVariable.value:
            question = Question(field)
            node = QuestionTreeItem(question)
            
            if parent is not None: 
                parent.add(question)
                
                if str(field.Parent.Parent.ObjectTypeValue) == objectTypeConstants.mtArray.value:
                    for v in field.Variables:
                        if ".." not in v.Indexes.split(','):
                            child_node = self.create_a_node(v, question)
                            node.addChild(child_node)
            return node
        elif str(field.ObjectTypeValue) == objectTypeConstants.mtRoutingItems.value:
            question = Question(field)
            if parent is not None: parent.add(question)
            node = QuestionTreeItem(question)
            return node
        else:
            question = Question(field)
            if parent is not None: parent.add(question)
            node = QuestionTreeItem(question)
            
            for f in field.Fields:
                child_node = self.create_a_node(f, question)
                node.addChild(child_node)

            return node

    def init_bvc_questions(self):
        f = open(r'temp\bvc_temp.json', mode = 'r', encoding="utf-8")
        bvc_variables = json.loads(f.read())
        f.close()
        
        self.tbl_bvc_questions.setColumnCount(3)
        self.tbl_bvc_questions.setRowCount(len(bvc_variables.keys()))
        self.tbl_bvc_questions.setHorizontalHeaderLabels(["Variable Name", "Variable Label", ""])
        
        self.tbl_bvc_questions.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl_bvc_questions.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl_bvc_questions.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)

        self.tbl_bvc_questions.setAcceptDrops(True)
        self.tbl_bvc_questions.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)

        self.tbl_bvc_questions.dragMoveEvent = self.dragMoveEvent
        self.tbl_bvc_questions.dragEnterEvent = self.dragEnterEvent
        self.tbl_bvc_questions.dropEvent = self.dropEvent

        i = 0

        for k, v in bvc_variables.items():
            for j in range(self.tbl_bvc_questions.columnCount()):
                item = QTableWidgetItem()

                if j == 0:
                    if v["properties"]["required_variable"]:
                        item.setCheckState(Qt.CheckState.Checked)
                        item.setFlags(~Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                    else:
                        item.setCheckState(Qt.CheckState.Unchecked)
                        item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                        
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)        
                    item.setText(k)

                    self.tbl_bvc_questions.setItem(i, j, item)
                elif j == 1:
                    item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    item.setText(v["label"])

                    self.tbl_bvc_questions.setItem(i, j, item)
                else:
                    if v["properties"]["allow_user_to_add_variables"]:
                        button = QPushButton("...")
                        button.setEnabled(False)
                        button.setToolTip("Add variables")
                        item.setSizeHint(button.sizeHint())
                        
                        self.tbl_bvc_questions.setItem(i, j, item)
                        self.tbl_bvc_questions.setCellWidget(i, j, button)
                    else:
                        item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                        self.tbl_bvc_questions.setItem(i, j, item)
            
            i = i + 1
        
        self.tbl_bvc_questions.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tbl_bvc_questions.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        
        self.tbl_bvc_questions.cellClicked.connect(self.handleCellClicked)
        self.tbl_bvc_questions.itemChanged.connect(self.handleItemChanged)

    def handleItemChanged(self, item):
        if item.column() == 0:
            checkbox = self.tbl_bvc_questions.item(item.row(), 0)
            
            if checkbox.checkState() == Qt.CheckState.Unchecked:
                self.tbl_bvc_questions.setItem(item.row(), 2, QTableWidgetItem(""))

    def handleCellClicked(self, row, col):
        item = self.tbl_bvc_questions.item(row, col)
        corresponding_item = self.tbl_bvc_questions.cellWidget(row, col + 2)

        if corresponding_item is not None: 
               
            corresponding_item.setEnabled(item.checkState() == Qt.CheckState.Checked)
            corresponding_item.clicked.connect(self.handleButtonClicked)

    def handleButtonClicked(self):
        #Determine which QPushButton triggered the event
        sender = self.sender()
        
        #determine which QTableWidgetItem contains the QPushButton 
        index = self.tbl_bvc_questions.indexAt(sender.pos())

        if index.isValid():
            item = self.tbl_bvc_questions.item(index.row(), index.column() - 2)

            self.variable_list_dialog = VariableListDialog()
            self.variable_list_dialog.setWindowTitle("List variables of {}".format(item.text()))

            if self.variable_list_dialog.exec():
                a = ""
    
    def hanldeItemPressed(self, event):
        if event.isSelected():
            self.ptxt_question_content.setPlainText(event.text(2))
            variables = event.question.get_variables_list()
            a = ""
            #data = QMimeData()
            #data.setText(event.text(1))
            
            

            #drag = QDrag(self.tree_questions)
            #drag.setMimeData(data)
            
            #drag.exec()

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            item = self.tbl_bvc_questions.itemAt(event.position().toPoint())
            
            if item:
                checkbox = self.tbl_bvc_questions.item(item.row(), 0)

                if checkbox.checkState() == Qt.CheckState.Checked:
                    event.accept()
                else:
                    event.ignore()
            else:
                event.ignore()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasText():
            item = self.tbl_bvc_questions.itemAt(event.position().toPoint())

            if item:
                checkbox = self.tbl_bvc_questions.item(item.row(), 0)

                if checkbox.checkState() == Qt.CheckState.Checked:
                    event.accept()
                else:
                    event.ignore()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasText():
            item = self.tbl_bvc_questions.itemAt(event.position().toPoint())
            
            self.tbl_bvc_questions.setItem(item.row(), 2, QTableWidgetItem(event.mimeData().text()))
            event.accept()