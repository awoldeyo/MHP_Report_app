import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from login import Ui_MainWindow
from mhp_jira import Connection


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.initUI()
        self.show()
        
    def initUI(self):
        self.ui.statusBar.showMessage("Not connected!")
        self.ui.pushButton.clicked.connect(self.establish_connection)
        self.ui.pushButton_2.clicked.connect(self.generate_reports)
    
    def establish_connection(self):
        self.username = self.ui.username.text()
        self.password = self.ui.password.text()
        self.connection = Connection(self.username, self.password)
        self.ui.statusBar.showMessage(self.connection.authenticate())
        
    def generate_reports(self):
        self.fileobject = self.saveFileDialog()
        if self.fileobject is not None:
            self.connection.create_excel(self.fileobject)
            QMessageBox.about(self, "Info", f"Successfully generated report!\n Path:{self.fileobject}")
        
    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"Save file as...","","Excel-Document (*.xlsx);;All Files (*)", options=options)
        if fileName:
            return fileName
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Window()
    w.show()
    sys.exit(app.exec_())