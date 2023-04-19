from PyQt6.QtWidgets import QMdiSubWindow, QMessageBox
from gui.frmCodeListing import Ui_code_listing_form
from objects.IOMObject import CodeListingWidgetItem

import json

class CodeListingForm(QMdiSubWindow, Ui_code_listing_form):
    def __init__(self):
        super().__init__()

        #load ui
        self.setupUi(self)

        self.code_list = dict()

        self.init()

        self.pbtn_create_new_code_list.clicked.connect(self.handle_create_new_codelist_clicked)
        self.pbtn_delete_code_list.clicked.connect(self.handle_delete_codelist_clicked)

    def init(self):
        self.wtree_code_list.clear()
        self.wtree_code_list.setHeaderLabel("Code Listing")

    def handle_create_new_codelist_clicked(self):
        self.code_list[self.line_code_list_name.text()] = list()

        with open("codelist.json", "w") as outfile:
            json_string = json.dumps(self.code_list, indent=4)
            outfile.write(json_string)

        node = CodeListingWidgetItem(self.line_code_list_name.text())
        self.wtree_code_list.addTopLevelItem(node)
        self.line_code_list_name.setText("")

    def handle_delete_codelist_clicked(self):
        selected_item = self.wtree_code_list.currentItem()
        index = self.wtree_code_list.indexOfTopLevelItem(selected_item)
        
        if selected_item is not None:
            result = QMessageBox.question(self, "Remove", "Are you sure you want to delete this code list '%s'?" % selected_item.text(0))

            if result == QMessageBox.StandardButton.Yes:
                self.wtree_code_list.takeTopLevelItem(index)
                
                