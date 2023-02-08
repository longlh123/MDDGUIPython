import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QMessageBox, QFileDialog
from gui.frmLogin import Ui_login_form
from main import MainWindow

class LoginForm(QWidget):
    def __init__(self):
        super().__init__()

        #use the Ui_login_form
        self.ui = Ui_login_form()
        self.ui.setupUi(self)

        #authenticate when the Login button is clicked
        self.ui.btn_login.clicked.connect(self.authenticate)

        #show the login window
        self.show()

    def authenticate(self):
        email = self.ui.edit_email_address.text()
        password = self.ui.edit_password.text()

        #validate the email address and password
        if email == 'a' and password == 'a':
            #open the main form if the login is successful
            self.main_form = MainWindow()
            self.main_form.showMaximized()

            #close the login form
            self.close()
        else:
            #show an error message if the login is unsuccessful
            QMessageBox.critical(self, "Error", "Invalid email or password.")
            
            self.ui.edit_email_address.setText("")
            self.ui.edit_password.setText("")
            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    login_window = LoginForm()
    sys.exit(app.exec())
