import win32com.client
import subprocess
import time
import sys
import pandas as pd

base_upload_simulador = r"D:\Users\bruno_cardozo\Downloads\Carga _Ivan_Cargaa.xlsx"
arquivo = pd.read_excel(base_upload_simulador)

user = "BCARDOZO"
password = "Longlive30@@"

class SapGui(object):
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(15)
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine

        self.connection = application.OpenConnection("2.01           Coty Brazil - SAP ECC - Development - Projects (ED7)", True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()

    def sapLogin(self):
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "130"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            self.session.FindById("wnd[0]").sendVKey(0)

        except:
            print(sys.exc_info()[0])


if __name__ == '__main__':
    SapGui().sapLogin()
