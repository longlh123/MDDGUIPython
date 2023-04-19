import os, sys
sys.path.append(os.getcwd())

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QHeaderView, QPushButton, QAbstractItemView
from PyQt6.QtCore import Qt, QFile, QIODevice, QXmlStreamReader
from PyQt6.QtGui import QStandardItemModel
from gui.frmMain import Ui_MainWindow

from dialogs.variable_list_dialog import VariableListDialog
from dialogs.query_tool_dialog import QueryToolDialog
from sub_windows.code_listing_form import CodeListingForm
from sub_windows.database_form import DatabaseMdiSubWindow
from objects.IOMObject import QuestionTreeItem

from pathlib import Path
import json
import win32com.client as w32

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, department):
        super().__init__()

        #Load the Ui
        self.setupUi(self) 

        self.mdd_path = ""
        self.department = department
        
        self.MDM = w32.Dispatch(r'MDM.Document')

        #Menu File
        self.actionOpen.triggered.connect(self.open_file_dialog)
        
        #Menu View
        self.actionProject_Box.toggled.connect(self.toggle_dock_widget)
        self.actionQuestion_Box.toggled.connect(self.toggle_dock_widget)
        self.actionQuery_Syntax_Box.toggled.connect(self.toggle_dock_widget)

        #Menu Variables
        self.actionShow_Open_Ended_Variables.toggled.connect(self.toggle_dock_widget)

        self.actionQuery_Tool.triggered.connect(self.handle_query_tool_triggered)

        self.dock_project.visibilityChanged.connect(self.handle_visibility_changed)
        self.dock_question_properties.visibilityChanged.connect(self.handle_visibility_changed)
        self.dock_query_syntax_box.visibilityChanged.connect(self.handle_visibility_changed)

        self.ptext_query_syntax.textChanged.connect(self.handle_query_syntax_textchanged)
        self.pbtn_apply.clicked.connect(self.handle_apply_query_syntax_clicked)

        #Menu Window
        self.actionDataBase.triggered.connect(self.window_triggered)
        self.actionCodeListing.triggered.connect(self.window_triggered)

        self.line_filter.returnPressed.connect(self.handle_return_pressed)

        self.init()

    def init(self, enabled=False):
        
        self.enable_actionWindow(enabled)
        self.enable_actionVariables(enabled)
        self.enable_actionBox(enabled)

        self.dock_project.setVisible(enabled)
        self.dock_question_properties.setVisible(enabled)
        self.dock_query_syntax_box.setVisible(enabled)

        self.close_all_subwindows()

    def enable_actionWindow(self, enable):
        self.actionDataBase.setEnabled(enable)
        self.actionCodeListing.setEnabled(enable)

    def enable_actionVariables(self, enable):
        self.actionShow_Open_Ended_Variables.setEnabled(enable)

    def enable_actionBox(self, enable):
        self.actionProject_Box.setEnabled(enable)
        self.actionQuestion_Box.setEnabled(enable)
        self.actionQuery_Syntax_Box.setEnabled(enable)

    def init_question_box(self):
        #load XML file
        file = QFile(r"view/temp/question_properties_table.xml")

        if not file.open(QIODevice.OpenModeFlag.ReadOnly | QIODevice.OpenModeFlag.Text):
            print("Failed to open XML file: ", file.errorString())
            sys.exit(1)
        
        #Create XML stream reader
        reader = QXmlStreamReader(file)

        #Read XML file
        while not reader.atEnd():
            if reader.isStartElement():
                if reader.name() == "TableSetting":
                    rowCount = int(reader.attributes().value("rows"))
                    columnCount = int(reader.attributes().value("columns"))

                    self.table_question_properties.setRowCount(rowCount)
                    self.table_question_properties.setColumnCount(columnCount)
                elif reader.name() == "Header":
                    while reader.readNextStartElement():
                        if reader.isStartElement():
                            print(reader.name(), "open")
                        reader.readNext()
                elif reader.name() == "Column":
                    print(reader.name())


            if reader.isEndElement():
                print(reader.name(), "close")
                """
                if reader.name() == "TableSetting":
                    rowCount = int(reader.attributes().value("rows"))
                    columnCount = int(reader.attributes().value("columns"))

                    self.table_question_properties.setRowCount(rowCount)
                    self.table_question_properties.setColumnCount(columnCount)
                elif reader.name() == "Header":                
                    col_names = list()

                    while not reader.isEndElement() or reader.name() != "Header":
                        if reader.isStartElement() and reader.name() == "Column":
                            col_names.append(reader.attributes().value("name"))
                        reader.readNext()

                    self.table_question_properties.setHorizontalHeaderLabels(col_names)
                elif reader.name() == "Rows":
                    while not reader.isEndElement() or reader.name() != "Rows":
                        if reader.isStartElement() and reader.name() == "Row":


                            print(reader.attributes().value("value"))
                        reader.readNext()
                """
            reader.readNext()  

        #Close file
        file.close()        

    def generate_categories_list(self, categories):
        self.tbl_categories_list.clear()
        self.tbl_categories_list.setRowCount(0)
        self.tbl_categories_list.setColumnCount(0)

        if categories is not None:
            self.tbl_categories_list.setColumnCount(2)
            self.tbl_categories_list.setRowCount(len(categories))
            self.tbl_categories_list.setHorizontalHeaderLabels(["Name", "Label"])

            self.tbl_categories_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            self.tbl_categories_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)

            row = 0

            for cat in categories:
                item = QTableWidgetItem()
                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                item.setText(cat.name)
                self.tbl_categories_list.setItem(row, 0, item)
                
                item = QTableWidgetItem()
                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                item.setText(cat.label)
                item.setToolTip(cat.label)
                self.tbl_categories_list.setItem(row, 1, item)

                row = row + 1

    def handle_query_tool_triggered(self):
        dialog = QueryToolDialog()

        if dialog.exec():
            a = ""

    def open_file_dialog(self):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setNameFilters(["Data Collection Data Files (*.mdd)", "Excel File (*.xlsx)"])
        dialog.setViewMode(QFileDialog.ViewMode.List)

        if dialog.exec():
            filenames = dialog.selectedFiles()

            if filenames:
                self.mdd_path = Path(filenames[0])

                QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
                
                self.enable_actionWindow(True)
                self.enable_actionVariables(True)
                self.enable_actionBox(True)
                
                self.init_questions()
                self.init_question_box()

                self.init(enabled=True)

                QApplication.restoreOverrideCursor()

    def handle_return_pressed(self):
        self.filter_variables(self.line_filter.text())

    def toggle_dock_widget(self, state):
        match self.sender().text():
            case "Project Box":
                self.dock_project.setVisible(state)
            case "Question Box":
                self.dock_question_properties.setVisible(state)
            case "Query Syntax Box":
                self.ptext_query_syntax.setPlainText("")
                self.dock_query_syntax_box.setVisible(state)
            case "Show Open-Ended Variables":
                self.line_filter.setText("")
                self.filter_variables(self.line_filter.text())

    def handle_query_syntax_textchanged(self):
        self.pbtn_apply.setEnabled(len(self.sender().toPlainText()) > 0)

    def handle_apply_query_syntax_clicked(self):
        query = self.ptext_query_syntax.toPlainText()



    def filter_variables(self, string=""):
        for i in range(self.wtree_questions.topLevelItemCount()):
            item = self.wtree_questions.topLevelItem(i)
            item.set_hidden(string, filter_open_ended_variables=self.actionShow_Open_Ended_Variables.isChecked())

    def handle_visibility_changed(self, visible):
        dock_panel = self.sender()

        match dock_panel.windowTitle():
            case "Project":
                self.actionProject_Box.setChecked(visible)
            case "Question":
                self.actionQuestion_Box.setChecked(visible)
            case "Query Syntax":
                self.actionQuery_Syntax_Box.setChecked(visible)

    def window_triggered(self, p):
        match self.sender().text():
            case "Database":
                sub = DatabaseMdiSubWindow()
            case "Code Listing":
                sub = CodeListingForm()
        
        if sub.windowTitle() not in [s.windowTitle() for s in self.mdiArea.subWindowList()]:
            self.mdiArea.addSubWindow(sub)
            sub.show()
            sub.closeEvent = lambda event: self.handle_sub_window_closed(sub, event)

    def handle_sub_window_closed(self, sub_window, event):
        for sub in self.mdiArea.subWindowList():
            if sub.windowTitle() == sub_window.windowTitle():
                self.mdiArea.removeSubWindow(sub)

    def close_all_subwindows(self):
        for sub in self.mdiArea.subWindowList():
            sub.close()

    def init_questions(self):
        try:
            self.MDM.Open(str(self.mdd_path))

            self.wtree_questions.clear()
            self.wtree_questions.setHeaderLabel(self.mdd_path.name)
            
            #Set the drap and drop in QTreeWidget and QTableWidget
            self.wtree_questions.setDragEnabled(True)
            self.wtree_questions.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
            
            for field in self.MDM.Fields:
                if field.Name == "S12b":
                    a = ""
                node = QuestionTreeItem(field)

                if node is not None:
                    self.wtree_questions.addTopLevelItem(node)
            
            self.MDM.Close()

            #Connect the 'itemPressed' signal to the 'handleStartDrap' method
            self.wtree_questions.itemPressed.connect(self.hanldeItemPressed)
        except AttributeError:
            a = ""
  
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
            variables = event.get_variables_list()
            #self.init_question_box(question_content=event.text(1), categories=event.question.categories)

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
