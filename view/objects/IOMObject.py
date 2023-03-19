from PyQt6.QtWidgets import QTreeWidgetItem
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QIcon

from objects.enumerations import dataTypeConstants, objectTypeConstants, objectDepartments

import pickle
import re

class Questions():
    def __init__(self):
        self.questions = list()

    def add(self, question):
        self.questions.append(question)

    def find(self, question):
        for qre in self.questions:
            if qre.Name == question.Name and qre.FullName == question.FullName:
                return qre
        return None

class Categories():
    def __init__(self):
        self.categories = list()

    def add(self, category):
        self.categories.append(category)

    def find(self, category):
        for cat in self.categories:
            if cat.Name == category.Name:
                return cat
        return None
    
class Question():
    def __init__(self, field):
        self.name = field.Name
        self.fullname = field.FullName
        self.objecttype = str(field.ObjectTypeValue)
        self.indexes = None if str(field.ObjectTypeValue) != objectTypeConstants.mtRoutingItems.value else field.Indexes
        self.set_label(field)
        self.icon = self.get_field_icon(field)
        self.questions = list()

    def set_label(self, field):
        if str(field.ObjectTypeValue) == objectTypeConstants.mtRoutingItems.value:
            self.label = self.get_variable_label(field.Parent.Parent, self.indexes)
        else:
            self.label = self.replace_label(field.Label) 

    def get_variable_label(self, field, category):
        if field.Parent.Parent is None:
            return field.Categories[re.sub(pattern="[\{\}]", repl="", string=category)].Label
        else:
            return "{} - {}".format(self.get_variable_label(field.Parent.Parent, self.indexes.split(',')[len(self.indexes.split(',')) - 2]), field.Categories[re.sub(pattern="[\{\}]", repl="", string=self.indexes.split(',')[len(self.indexes.split(',')) - 1])].Label) 

    def add(self, question):
        self.questions.append(question)

    def get_field_icon(self, field):
        root = 'view/images/questions'
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
                case dataTypeConstants.mtNone.value:
                    image_name = 'Display.png'
        else:
            match str(field.ObjectTypeValue):
                case objectTypeConstants.mtClass.value:
                    image_name = 'Block.png'
                case objectTypeConstants.mtArray.value:
                    types = list()
                    
                    try:
                        for item in field.Items:
                            if item.DataType not in types:
                                if item.DataType == dataTypeConstants.mtCategorical.value:
                                    if item.MinValue == 1 and item.MaxValue == 1:
                                        types.append(31)
                                    else:
                                        types.append(32)
                                else:
                                    types.append(item.DataType)
                        
                        if len(types) == 1:
                            match field.Items[0].DataType:
                                case dataTypeConstants.mtBoolean.value:
                                    image_name = 'Grid.png'
                                case dataTypeConstants.mtCategorical.value:
                                    if field.Items[0].MinValue == 1 and field.Items[0].MaxValue == 1:
                                        image_name = 'SingleResponseGrid.png'
                                    else:
                                        image_name = 'MultipeResponseGrid.png'
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
                    except AttributeError:
                        image_name = 'Loop.png'

        return "{}/{}".format(root, image_name)
    
    def replace_label(self, label):
        blacklist = ["SHOWTABLET", "SHOW TABLET", "SHOWTABLET THANG ĐIỂM", "SHOWPHOTO"]
        s = re.sub(pattern=".*(?=({}))".format("|".join(blacklist)), repl="", string=label)
        s = re.sub(pattern="({})".format("|".join(blacklist)), repl="", string=s)
        return s 

    def get_variables_list(self, question=None):
        variables_list = list()

        if question is None:
            if len(self.questions) == 0:
                variables_list.append(self.fullname)
            else:
                for q in self.questions:
                    variables_list.extend(self.get_variables_list(q))
        else:
            if len(question.questions) == 0:
                variables_list.append(question.fullname)
            else:
                for q in question.questions:
                    variables_list.extend(self.get_variables_list(q))
        
        return variables_list



class QuestionTreeItem(QTreeWidgetItem):
    def __init__(self, question):
        super().__init__(["", question.fullname, question.label])
        self.question = question
        self.setIcon(0, QIcon(self.question.icon))
        self.set_text()

    def set_text(self):
        match self.question.objecttype:
            case objectTypeConstants.mtVariable.value | objectTypeConstants.mtArray.value | objectTypeConstants.mtClass.value:
                self.setText(0, self.question.name)
            case objectTypeConstants.mtRoutingItems.value:
                self.setText(0, self.question.indexes)

    def serialize_question(self):
        return pickle.dumps(self.question)

    def set_mime_data(self):
        mime_data = QMimeData()
        mime_data.setData("application/x-question", self.serialize_question())
        self.setData(0, Qt.ItemDataRole.UserRole, mime_data)
    
#class Category(Field):
#    def __init__(self, category):
#        Field.__init__(self, category)
#        self.categories = Categories()

