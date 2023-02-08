import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QHeaderView, QPushButton, QTreeWidgetItem, QAbstractItemView
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QIcon, QDrag
from gui.frmMain import Ui_MainWindow

from dialogs.variable_list_dialog import VariableListDialog

from pathlib import Path
import json
import win32com.client as w32
from enumerations import dataTypeConstants, objectTypeConstants

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()

        #Load the Ui
        self.setupUi(self) 

        self.mdd_path = ""

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
                self.init_bvc_questions()
                
                QApplication.restoreOverrideCursor()

    def init_questions(self):
        self.MDM.Open(str(self.mdd_path))

        self.tree_questions.clear()
        self.tree_questions.setHeaderLabel(self.mdd_path.name)
        
        #Set the drap and drop in QTreeWidget and QTableWidget
        self.tree_questions.setDragEnabled(True)
        
        #Connect the 'itemPressed' signal to the 'handleStartDrap' method
        self.tree_questions.itemPressed.connect(self.hanldeItemPressed)

        for field in self.MDM.Fields:
            self.add_node(field, parent_node=self.tree_questions)

        self.MDM.Close()

    def add_node(self, field, parent_node=None, variables=list()):
        if str(field.ObjectTypeValue) == objectTypeConstants.mtVariable.value:
            node = QTreeWidgetItem(parent_node)
            node.setText(0, field.Name)
            node.setIcon(0, self.get_field_icon(field))

            if field.DataType == dataTypeConstants.mtCategorical.value:
                if field.OtherCategories.Count > 0:
                    for helperfield in field.HelperFields:
                        self.add_node(helperfield, parent_node=node)
            
            if len(variables) > 0:
                for variable in variables:
                    self.add_node(variable, parent_node=node)
        elif str(field.ObjectTypeValue) == objectTypeConstants.mtRoutingItems.value:
            node = QTreeWidgetItem(parent_node)
            node.setText(0, field.Indexes)
            node.setIcon(0, self.get_field_icon(field))
        else:
            p_node = QTreeWidgetItem(parent_node)
            p_node.setText(0, field.Name)
            
            child_nodes = list()

            for f in field.Fields:
                if str(field.ObjectTypeValue) == objectTypeConstants.mtArray.value:
                    self.add_node(f, parent_node=p_node, variables=f.Variables)

                    if str(f.ObjectTypeValue) == objectTypeConstants.mtVariable.value:
                        if f not in child_nodes:
                            child_nodes.append(f)
                else:
                    self.add_node(f, parent_node=p_node)
            
            p_node.setIcon(0, self.get_field_icon(field, child_nodes=child_nodes))
    
    def get_field_icon(self, field, child_nodes=list()):
        root = 'images/questions'
        image_name = ''

        if str(field.ObjectTypeValue) == objectTypeConstants.mtVariable.value or str(field.ObjectTypeValue) == objectTypeConstants.mtRoutingItems.value:
            match field.DataType:
                case dataTypeConstants.mtBoolean.value:
                    image_name = 'Boolean.png'
                case dataTypeConstants.mtCategorical.value:
                    if field.MinValue == 1 and field.MaxValue == 1:
                        image_name = 'SingleResponse.png'
                    else:
                        image_name = 'MultipleResponse.png'
                case dataTypeConstants.mtDate.value:
                    image_name = 'DateTime.png'
                case dataTypeConstants.mtDouble.value:
                    image_name = 'Numeric.png'
                case dataTypeConstants.mtLong.value:
                    image_name = 'Numeric.png'
                case dataTypeConstants.mtText.value:
                    image_name = 'Text.png'
                case dataTypeConstants.mtNone:
                    image_name = 'Display.png'
        else:
            match str(field.ObjectTypeValue):
                case objectTypeConstants.mtClass.value:
                    image_name = 'Block.png'
                case objectTypeConstants.mtArray.value:
                    if len(child_nodes) == 1:
                        match child_nodes[0].DataType:
                            case dataTypeConstants.mtBoolean.value:
                                image_name = 'Grid.png'
                            case dataTypeConstants.mtCategorical.value:
                                if child_nodes[0].MinValue == 1 and child_nodes[0].MaxValue == 1:
                                    image_name = 'SingleResponseGrid.png'
                                else:
                                    image_name = 'NumericResponseGrid.png'
                            case dataTypeConstants.mtDate.value:
                                image_name = 'DateTime.png'
                            case dataTypeConstants.mtDouble.value:
                                image_name = 'NumericResponseGrid.png'
                            case dataTypeConstants.mtLong.value:
                                image_name = 'NumericResponseGrid.png'
                            case dataTypeConstants.mtText.value:
                                image_name = 'TextResponseGrid.png'
                    else:
                        image_name = 'Loop.png'

        return QIcon("{}/{}".format(root, image_name))

    def init_bvc_questions(self):
        f = open(r'temp\bvc_temp.json', mode = 'r', encoding="utf-8")
        bvc_variables = json.loads(f.read())
        f.close()
        
        self.tbl_bvc_questions.setColumnCount(3)
        self.tbl_bvc_questions.setRowCount(len(bvc_variables.keys()))
        self.tbl_bvc_questions.setHorizontalHeaderLabels(["Variable Name", "Variable Label", ""])

        self.tbl_bvc_questions.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl_bvc_questions.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl_bvc_questions.setColumnWidth(2, 20)

        self.tbl_bvc_questions.setAcceptDrops(True)
        self.tbl_bvc_questions.setDragDropOverwriteMode(False)
        self.tbl_bvc_questions.dragEnterEvent = self.dragEnterEvent
        self.tbl_bvc_questions.dropEvent = self.dropEvent

        i = 0

        for k, v in bvc_variables.items():
            for j in range(self.tbl_bvc_questions.columnCount()):
                item = QTableWidgetItem()

                if j == 0:
                    if v["properties"]["required_variable"]:
                        item.setCheckState(Qt.CheckState.Checked)
                        item.setFlags(~Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                    else:
                        item.setCheckState(Qt.CheckState.Unchecked)
                        item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                    
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)        
                    item.setText(k)

                    self.tbl_bvc_questions.setItem(i, j, item)
                elif j == 1:
                    item.setFlags(Qt.ItemFlag.ItemIsEnabled)
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
                        item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                        self.tbl_bvc_questions.setItem(i, j, item)
            
            i = i + 1
        
        self.tbl_bvc_questions.cellClicked.connect(self.handleCellClicked)
        
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
       
    def hanldeItemPressed(self, item):
        if item.isSelected():
            drag = QDrag(self.tree_questions)
            data = QMimeData()
            data.setText(item.text(0))
            drag.setMimeData(data)
            drag.exec()

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        self.tbl_bvc_questions.setItem(0, 0, QTableWidgetItem(event.mimeData().text()))
        event.acceptProposedAction()
        #mime_data = event.mimeData()