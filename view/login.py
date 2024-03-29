import os, sys
sys.path.append(os.getcwd())

from PyQt6.QtWidgets import QApplication, QWidget, QMessageBox
from gui.frmLogin import Ui_login_form
from main import MainWindow

from objects.enumerations import objectDepartments
from controller.LoginController import LoginController 

class LoginForm(QWidget):
    def __init__(self):
        super().__init__()

        #use the Ui_login_form
        self.ui = Ui_login_form()
        self.ui.setupUi(self)

        self.loginController = LoginController()

        #authenticate when the Login button is clicked
        self.ui.btn_login.clicked.connect(self.authenticate)

        #show the login window
        self.show()

    def authenticate(self): 
        email = self.ui.edit_email_address.text()
        password = self.ui.edit_password.text()
    
        #validate the email address and password
        if self.loginController.authenticate(email, password):
            #open the main form if the login is successful
            self.main_form = MainWindow(objectDepartments.DP.value)
            self.main_form.showMaximized()
            
            #close the login form
            self.close()
        else:
            #show an error message if the login is unsuccessful
            QMessageBox.critical(self, "Error", "Invalid email or password.", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            
            self.ui.edit_email_address.setText("")
            self.ui.edit_password.setText("")
            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    login_window = LoginForm()
    sys.exit(app.exec())
