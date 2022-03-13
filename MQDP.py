# -*- coding: utf-8 -*- 

import os
import sys

from PyQt5 import (QtCore, QtGui)
from PyQt5.QtWidgets import (QWidget, QLabel, QTextEdit, QLineEdit, QPushButton,
    QFrame, QApplication, QMessageBox, QGridLayout, QComboBox, QFileDialog)


class MainWidget(QWidget):

    __grid = None

    __standardsCombo = None

    __pathLineEdit = None
    __pathInteractButton = None

    __outFolderLineEdit = None
    __outFolderInteractButton = None
    __outCombo = None
    __outs = ["GIFT media"]

    __startButton = None

    __allStandards = ["standardK"]
    __curStandard = None # str

    def __init__(self):
        super().__init__()

        self.__initUI()


    def __initUI(self):

        self.__grid = QGridLayout()
        self.__grid.setSpacing(10)

        standardsLbl = QLabel("Standards:", self)

        self.__standardsCombo = QComboBox(self)
        self.__standardsCombo.addItems(self.__allStandards)
        self.__standardsCombo.activated[str].connect(self.__standardComboActivated)

        #================================================
        pathLbl = QLabel("Path to docx file:", self)

        self.__pathLineEdit = QLineEdit(self)

        self.__pathInteractButton = QPushButton("Choose docx file", self)
        self.__pathInteractButton.clicked.connect(lambda:self.__pathInteractButton_hundler(self.__pathInteractButton))

        #================================================
        outsLbl = QLabel("Out:", self)

        self.__outFolderLineEdit = QLineEdit(self)

        self.__outFolderInteractButton = QPushButton("Choose out folder", self)
        self.__outFolderInteractButton.clicked.connect(lambda:self.__outFolderInteractButton_hundler(self.__outFolderInteractButton))

        self.__outCombo = QComboBox(self)
        self.__outCombo.addItems(self.__outs)

        #================================================
        self.__startButton = QPushButton("Start")
        self.__startButton.clicked.connect(lambda:self.__startButton_handler(self.__startButton))

        #================================================
        self.__grid.addWidget(standardsLbl, 0, 0, 1, 1)
        self.__grid.addWidget(self.__standardsCombo, 1, 0, 1, 1)
        self.__grid.addWidget(QLabel("", self), 2, 0, 1, 1) # ===
        self.__grid.addWidget(pathLbl, 3, 0, 1, 1)
        self.__grid.addWidget(self.__pathLineEdit, 4, 0, 1, 1)
        self.__grid.addWidget(self.__pathInteractButton, 5, 0, 1, 1)
        self.__grid.addWidget(QLabel("", self), 6, 0, 1, 1) # ===
        self.__grid.addWidget(outsLbl, 7, 0, 1, 1)
        self.__grid.addWidget(self.__outFolderLineEdit, 8, 0, 1, 1)
        self.__grid.addWidget(self.__outFolderInteractButton, 9, 0, 1, 1)
        self.__grid.addWidget(self.__outCombo, 10, 0, 1, 1)
        self.__grid.addWidget(QLabel("", self), 11, 0, 1, 1) # ===
        self.__grid.addWidget(self.__startButton, 12, 0, 1, 1)

        self.setLayout(self.__grid)

        self.show()

    def __standardComboActivated(self, text):
        #self.__curStandard = text
        pass

    def __pathInteractButton_hundler(self, b):
        #filepath = QtWidgets.QFileDialog.getOpenFileName(self, 'Select docx file')
        curdir = str(os.getcwd()) # working directory
        filepath = QFileDialog.getOpenFileName(self, 'Select docx file', curdir, "Docx document (*.docx)")[0]
        self.__pathLineEdit.setText(filepath)

        if(self.__outFolderLineEdit.text() == ""):
            #outPath = os.path.dirname(os.path.abspath(filepath))
            outPath = filepath[:filepath.rfind(".docx")]
            self.__outFolderLineEdit.setText(outPath)


    def __outFolderInteractButton_hundler(self, b):
        curdir = str(os.getcwd()) # working directory
        filepath = QFileDialog.getExistingDirectory(self, 'Select out folder', curdir)
        self.__outFolderLineEdit.setText(filepath)

    def __startButton_handler(self, b):
        self.__curStandard = self.__standardsCombo.currentText()
        if(self.__curStandard == None):
            self.__ifError("Please, select standard", 4)
            return
        
        path = self.__pathLineEdit.text()
        if(path == None or path == ""):
            self.__ifError("Please, set path to docx file", 4)
            return

        #outPath = os.path.dirname(os.path.abspath(path))

        outPath = self.__outFolderLineEdit.text()
        if(outPath == None or outPath == ""):
            self.__ifError("Please, select out folder", 4)
            return

        if not os.path.exists(outPath):
            os.makedirs(outPath)

        if os.path.isdir(outPath):
            if not os.listdir(outPath):
                pass
            else:
                self.__ifError("Out folder must be empty", 4)
                return
        else:
            self.__ifError("Out folder does not exists", 4)
            return

        if(self.__curStandard == self.__allStandards[0]):
            from MQPD_standards import standardk_run
            res = standardk_run(path, outPath)
            self.__ifError(res[0] + "\n\n" + res[2], res[1])
        else:
            self.__ifError("No that standard", 4)
            return

    '''
    0 - None
    1 - Question
    2 - Information
    3 - Warning
    4 - Critical
    '''
    def __ifError(self, text : str, type: int):
        msg = QMessageBox()
        suffix = ""

        if(type == 0):
            msg.setWindowTitle("")
        elif(type == 1):
            msg.setIcon(QMessageBox.Question)
            msg.setWindowTitle("Question")
        elif(type == 2):
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Info")
        elif(type == 3):
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Warning")
        elif(type == 4):
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
            suffix = "Check README.md: https://github.com/The220th/MQDP/README.md"

        msg.setText(text + f"\n{suffix}")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setModal(True)
        msg.exec()






if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWidget = MainWidget()
    mainWidget.setWindowTitle("MQDP")
    mainWidget.setWindowIcon( QtGui.QIcon("./imgsrc/icon.svg") )
    sys.exit(app.exec_())