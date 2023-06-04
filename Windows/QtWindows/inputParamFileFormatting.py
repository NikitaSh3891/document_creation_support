# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'inputParamFileFormatting.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QCheckBox

from PyQt5.QtWidgets import QFileDialog

from WorkWithDB import findFileExtension
from Windows.QtWindows.errWindow import showErrWindow, editText


class Ui_inputParamFileFormattingWindow(object):
    def setupUi(self, inputParamFileFormattingWindow):
        inputParamFileFormattingWindow.setObjectName("inputParamFileFormattingWindow")
        inputParamFileFormattingWindow.setWindowModality(QtCore.Qt.NonModal)
        inputParamFileFormattingWindow.resize(800, 600)
        font = QtGui.QFont()
        font.setBold(False)
        font.setWeight(50)
        inputParamFileFormattingWindow.setFont(font)
        inputParamFileFormattingWindow.setStyleSheet("background-color: rgb(212, 237, 255);")
        self.centralwidget = QtWidgets.QWidget(inputParamFileFormattingWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.hiLabel = QtWidgets.QLabel(self.centralwidget)
        self.hiLabel.setGeometry(QtCore.QRect(0, 0, 800, 100))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(20)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.hiLabel.setFont(font)
        self.hiLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.hiLabel.setAutoFillBackground(False)
        self.hiLabel.setStyleSheet("background-color: rgb(180, 197, 255);")
        self.hiLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.hiLabel.setWordWrap(True)
        self.hiLabel.setObjectName("hiLabel")
        self.continueButton = QtWidgets.QPushButton(self.centralwidget)
        self.continueButton.setGeometry(QtCore.QRect(300, 540, 200, 50))
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
        self.infoLabel.setGeometry(QtCore.QRect(10, 110, 800, 60))
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
        self.infoLabel.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.infoLabel.setWordWrap(False)
        self.infoLabel.setObjectName("infoLabel")
        self.infoLabel_2 = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel_2.setGeometry(QtCore.QRect(10, 230, 541, 60))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.infoLabel_2.setFont(font)
        self.infoLabel_2.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.infoLabel_2.setAutoFillBackground(False)
        self.infoLabel_2.setStyleSheet("")
        self.infoLabel_2.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.infoLabel_2.setWordWrap(False)
        self.infoLabel_2.setObjectName("infoLabel_2")
        self.editFileFormattingCheckBox = QtWidgets.QCheckBox(self.centralwidget)
        self.editFileFormattingCheckBox.setGeometry(QtCore.QRect(580, 240, 31, 41))
        self.editFileFormattingCheckBox.setStyleSheet("QCheckBox {\n"
            "    spacing: 5px;\n"
            "    font-size:25px;     \n"
            "}\n"
            "\n"
            "QCheckBox::indicator {\n"
            "    width:  25px;\n"
            "    height: 25px;\n"
            "}")
        self.editFileFormattingCheckBox.setText("")
        self.editFileFormattingCheckBox.setObjectName("editFileFormattingCheckBox")
        self.infoLabel_3 = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel_3.setGeometry(QtCore.QRect(10, 300, 361, 60))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.infoLabel_3.setFont(font)
        self.infoLabel_3.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.infoLabel_3.setAutoFillBackground(False)
        self.infoLabel_3.setStyleSheet("")
        self.infoLabel_3.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.infoLabel_3.setWordWrap(False)
        self.infoLabel_3.setObjectName("infoLabel_3")
        self.createStatByFormattingCheckBox = QtWidgets.QCheckBox(self.centralwidget)
        self.createStatByFormattingCheckBox.setGeometry(QtCore.QRect(580, 310, 31, 31))
        self.createStatByFormattingCheckBox.setStyleSheet("QCheckBox {\n"
            "    spacing: 5px;\n"
            "    font-size:25px;     \n"
            "}\n"
            "\n"
            "QCheckBox::indicator {\n"
            "    width:  25px;\n"
            "    height: 25px;\n"
            "}")
        self.createStatByFormattingCheckBox.setText("")
        self.createStatByFormattingCheckBox.setObjectName("createStatByFormattingCheckBox")
        self.inputFilesButton = QtWidgets.QPushButton(self.centralwidget)
        self.inputFilesButton.setGeometry(QtCore.QRect(520, 430, 151, 41))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.inputFilesButton.setFont(font)
        self.inputFilesButton.setStyleSheet("background-color: rgb(180, 197, 255);")
        self.inputFilesButton.setDefault(False)
        self.inputFilesButton.setFlat(False)
        self.inputFilesButton.setObjectName("inputFilesButton")
        self.fileNamesPlainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.fileNamesPlainTextEdit.setGeometry(QtCore.QRect(10, 380, 491, 141))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(12)
        self.fileNamesPlainTextEdit.setFont(font)
        self.fileNamesPlainTextEdit.setObjectName("fileNamesPlainTextEdit")
        self.infoLabel_4 = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel_4.setGeometry(QtCore.QRect(10, 170, 330, 60))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.infoLabel_4.setFont(font)
        self.infoLabel_4.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.infoLabel_4.setAutoFillBackground(False)
        self.infoLabel_4.setStyleSheet("")
        self.infoLabel_4.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.infoLabel_4.setWordWrap(False)
        self.infoLabel_4.setObjectName("infoLabel_4")
        self.fileFormattingComboBox = QtWidgets.QComboBox(self.centralwidget)
        self.fileFormattingComboBox.setGeometry(QtCore.QRect(415, 187, 350, 30))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        self.fileFormattingComboBox.setFont(font)
        self.fileFormattingComboBox.setObjectName("fileFormattingComboBox")
        inputParamFileFormattingWindow.setCentralWidget(self.centralwidget)

        self.fileNamesPlainTextEdit.setReadOnly(True)

        from WorkWithDB import findFileFormattingPattern
        patternArr = findFileFormattingPattern()
        for i in range(len(patternArr)):
            self.fileFormattingComboBox.addItem(patternArr[i][1])

        def inputFiles():
            self.fileNamesPlainTextEdit.setPlainText('')
            fileNames = QFileDialog.getOpenFileNames(None, "Выберете файл", "", "All Files (*)")[0]
            fileExtensions = findFileExtension()
            for fName in fileNames:
                isTrue = True
                for extension in fileExtensions:
                    if fName[fName.rfind(".") + 1:] == extension[1]:
                        isTrue = False
                        break
                if isTrue:
                    editText(f'Ошибка! Выбранный формат файла "{fName[fName.find("."):]}" не поддерживается')
                    showErrWindow()
                else:
                    self.fileNamesPlainTextEdit.appendHtml(fName)

        self.inputFilesButton.clicked.connect(inputFiles)

        def clickedContinueButton():
            if len(self.fileNamesPlainTextEdit.toPlainText()) != 0:
                idPattern = 0
                for i in range(len(patternArr)):
                    if self.fileFormattingComboBox.currentText() == patternArr[i][1]:
                        idPattern = patternArr[i][0]
                from Windows.QtWindows.tableParamFormattingWindow import showTableParamFormattingWindow
                showTableParamFormattingWindow()
                closeInputParamFileFormattingWindow()
                from InputTest import editTypeFileFormatting
                editTypeFileFormatting(self.editFileFormattingCheckBox.isChecked(),
                                       self.createStatByFormattingCheckBox.isChecked(),
                                       self.fileNamesPlainTextEdit.toPlainText(),
                                       idPattern)
            else:
                editText("Ошибка! Вы не выбрали ни одного файла для загрузки")
                showErrWindow()

        self.continueButton.clicked.connect(clickedContinueButton)

        self.retranslateUi(inputParamFileFormattingWindow)
        QtCore.QMetaObject.connectSlotsByName(inputParamFileFormattingWindow)

    def retranslateUi(self, inputParamFileFormattingWindow):
        _translate = QtCore.QCoreApplication.translate
        inputParamFileFormattingWindow.setWindowTitle(_translate("inputParamFileFormattingWindow", "Experiment"))
        self.hiLabel.setText(_translate("inputParamFileFormattingWindow", "Форматирование документов"))
        self.continueButton.setText(_translate("inputParamFileFormattingWindow", "Продолжить"))
        self.infoLabel.setText(_translate("inputParamFileFormattingWindow", "Выберите параметры для форматирования:"))
        self.infoLabel_2.setText(_translate("inputParamFileFormattingWindow", "Создать исправленный вариант документа(ов)?"))
        self.infoLabel_3.setText(_translate("inputParamFileFormattingWindow", "Создать файл со статистикой?"))
        self.inputFilesButton.setText(_translate("inputParamFileFormattingWindow", "Выбрать файлы"))
        self.infoLabel_4.setText(_translate("inputParamFileFormattingWindow", "Выберете форматирование:"))


def showInputParamFileFormattingWindow():
    global inputParamFileFormattingWindow
    inputParamFileFormattingWindow = QtWidgets.QMainWindow()
    ui = Ui_inputParamFileFormattingWindow()
    ui.setupUi(inputParamFileFormattingWindow)
    inputParamFileFormattingWindow.show()


def closeInputParamFileFormattingWindow():
    inputParamFileFormattingWindow.close()


def editIdUser(_idUser):
    global idUser
    idUser = _idUser
