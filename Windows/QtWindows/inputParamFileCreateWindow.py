# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'inputParamFileCreateWindow.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog


class Ui_inputParamFileCreateWindow(object):
    def setupUi(self, inputParamFileCreateWindow):
        inputParamFileCreateWindow.setObjectName("inputParamFileCreateWindow")
        inputParamFileCreateWindow.setWindowModality(QtCore.Qt.NonModal)
        inputParamFileCreateWindow.resize(800, 600)
        font = QtGui.QFont()
        font.setBold(False)
        font.setWeight(50)
        inputParamFileCreateWindow.setFont(font)
        inputParamFileCreateWindow.setStyleSheet("background-color: rgb(212, 237, 255);")
        self.centralwidget = QtWidgets.QWidget(inputParamFileCreateWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(0, 0, 800, 100))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(22)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label.setAutoFillBackground(False)
        self.label.setStyleSheet("background-color: rgb(180, 197, 255);")
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setWordWrap(False)
        self.label.setObjectName("label")
        self.continueButton = QtWidgets.QPushButton(self.centralwidget)
        self.continueButton.setGeometry(QtCore.QRect(300, 500, 200, 50))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.continueButton.setFont(font)
        self.continueButton.setStyleSheet("background-color: rgb(180, 197, 255);")
        self.continueButton.setDefault(False)
        self.continueButton.setFlat(False)
        self.continueButton.setObjectName("continueButton")
        self.infoLabel = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel.setGeometry(QtCore.QRect(0, 100, 800, 60))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(18)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.infoLabel.setFont(font)
        self.infoLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.infoLabel.setAutoFillBackground(False)
        self.infoLabel.setStyleSheet("")
        self.infoLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.infoLabel.setWordWrap(False)
        self.infoLabel.setObjectName("infoLabel")
        self.fileNameLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.fileNameLineEdit.setGeometry(QtCore.QRect(160, 170, 500, 45))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        self.fileNameLineEdit.setFont(font)
        self.fileNameLineEdit.setObjectName("fileNameLineEdit")
        self.infoLabel_2 = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel_2.setGeometry(QtCore.QRect(0, 220, 800, 60))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(18)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.infoLabel_2.setFont(font)
        self.infoLabel_2.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.infoLabel_2.setAutoFillBackground(False)
        self.infoLabel_2.setStyleSheet("")
        self.infoLabel_2.setAlignment(QtCore.Qt.AlignCenter)
        self.infoLabel_2.setWordWrap(False)
        self.infoLabel_2.setObjectName("infoLabel_2")
        self.fileFormattingComboBox = QtWidgets.QComboBox(self.centralwidget)
        self.fileFormattingComboBox.setGeometry(QtCore.QRect(200, 290, 400, 50))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        self.fileFormattingComboBox.setFont(font)
        self.fileFormattingComboBox.setObjectName("fileFormattingComboBox")
        self.infoLabel_3 = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel_3.setGeometry(QtCore.QRect(0, 360, 800, 60))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(18)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.infoLabel_3.setFont(font)
        self.infoLabel_3.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.infoLabel_3.setAutoFillBackground(False)
        self.infoLabel_3.setStyleSheet("")
        self.infoLabel_3.setAlignment(QtCore.Qt.AlignCenter)
        self.infoLabel_3.setWordWrap(False)
        self.infoLabel_3.setObjectName("infoLabel_3")
        self.titlePageComboBox = QtWidgets.QComboBox(self.centralwidget)
        self.titlePageComboBox.setGeometry(QtCore.QRect(200, 430, 400, 50))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        self.titlePageComboBox.setFont(font)
        self.titlePageComboBox.setObjectName("fileFormattingComboBox_2")
        inputParamFileCreateWindow.setCentralWidget(self.centralwidget)

        from WorkWithDB import findFileFormattingPattern
        patternArr = findFileFormattingPattern()
        for i in range(len(patternArr)):
            self.fileFormattingComboBox.addItem(patternArr[i][1])

        from WorkWithDB import findTitlePage
        titlePageArr = findTitlePage()
        self.titlePageComboBox.addItem("Без титульного листа")
        for i in range(len(titlePageArr)):
            self.titlePageComboBox.addItem(titlePageArr[i][1])

        def continueCreateFile():
            idPattern = 0
            for i in range(len(patternArr)):
                if self.fileFormattingComboBox.currentText() == patternArr[i][1]:
                    idPattern = patternArr[i][0]
            idTitlePage = 0
            if self.titlePageComboBox.currentText() == "Без титульного листа":
                idTitlePage = -1
            else:
                for i in range(len(titlePageArr)):
                    if self.titlePageComboBox.currentText() == titlePageArr[i][1]:
                        idTitlePage = titlePageArr[i][0]

            from InputTest import fileNameTest, editTitlePage, editFileFormatting, editIsUploadFile
            editIsUploadFile(False)
            editFileFormatting(idPattern)
            editTitlePage(idTitlePage)
            fileNameTest(self.fileNameLineEdit.text())

        self.continueButton.clicked.connect(continueCreateFile)

        self.retranslateUi(inputParamFileCreateWindow)
        QtCore.QMetaObject.connectSlotsByName(inputParamFileCreateWindow)

    def retranslateUi(self, inputParamFileCreateWindow):
        _translate = QtCore.QCoreApplication.translate
        inputParamFileCreateWindow.setWindowTitle(_translate("inputParamFileCreateWindow", "Experiment"))
        self.label.setText(_translate("inputParamFileCreateWindow", "Создание документа"))
        self.continueButton.setText(_translate("inputParamFileCreateWindow", "Продолжить"))
        self.infoLabel.setText(_translate("inputParamFileCreateWindow", "Ведите имя создаваемого файла:"))
        self.fileNameLineEdit.setPlaceholderText(_translate("inputParamFileCreateWindow", "Пример: example"))
        self.infoLabel_2.setText(_translate("inputParamFileCreateWindow", "Выберете форматирование создаваемого файла:"))
        self.infoLabel_3.setText(_translate("inputParamFileCreateWindow", "Выберете титульный лист создаваемого файла:"))


def showInputParamFileCreateWindow():
    global inputParamFileCreateWindow
    inputParamFileCreateWindow = QtWidgets.QMainWindow()
    ui = Ui_inputParamFileCreateWindow()
    ui.setupUi(inputParamFileCreateWindow)
    inputParamFileCreateWindow.show()


def closeInputParamFileCreateWindow():
    inputParamFileCreateWindow.close()