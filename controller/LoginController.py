
import os, sys
sys.path.append(os.getcwd())

from model.LoginModel import LoginModel

class LoginController:
    def __init__(self):
        self.loginModel = LoginModel()

    def authenticate(self, username, password):
        return self.loginModel.authenticate(username, password)

    