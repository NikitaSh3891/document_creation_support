# -*- coding: utf-8 -*-


# Form implementation generated from reading ui file 'createFileStyleWindow.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.
import re

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from Windows.QtWindows.errWindow import editText
from WorkWithDB import findIdPatternByName, createNewPattern, fillCollectionInPattern, fillFormattingUnit
class Ui_createFileStyleWindow(object):
    def setupUi(self, createFileStyleWindow):
        createFileStyleWindow.setObjectName("createFileStyleWindow")
        createFileStyleWindow.setWindowModality(QtCore.Qt.NonModal)
        createFileStyleWindow.resize(800, 600)
        font = QtGui.QFont()
        font.setBold(False)
        font.setWeight(50)
        createFileStyleWindow.setFont(font)
        createFileStyleWindow.setStyleSheet("background-color: rgb(212, 237, 255);")
        self.centralwidget = QtWidgets.QWidget(createFileStyleWindow)
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
        self.infoLabel = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel.setGeometry(QtCore.QRect(10, 120, 771, 101))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.infoLabel.setFont(font)
        self.infoLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.infoLabel.setAutoFillBackground(False)
        self.infoLabel.setStyleSheet("")
        self.infoLabel.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.infoLabel.setWordWrap(True)
        self.infoLabel.setObjectName("infoLabel")
        self.inputFilesButton = QtWidgets.QPushButton(self.centralwidget)
        self.inputFilesButton.setGeometry(QtCore.QRect(530, 230, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.inputFilesButton.setFont(font)
        self.inputFilesButton.setStyleSheet("background-color: rgb(180, 197, 255);")
        self.inputFilesButton.setDefault(False)
        self.inputFilesButton.setFlat(False)
        self.inputFilesButton.setObjectName("inputFilesButton")
        self.temporaryLabel = QtWidgets.QLabel(self.centralwidget)
        self.temporaryLabel.setGeometry(QtCore.QRect(10, 240, 500, 41))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.temporaryLabel.setFont(font)
        self.temporaryLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.temporaryLabel.setAutoFillBackground(False)
        self.temporaryLabel.setStyleSheet("")
        self.temporaryLabel.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.temporaryLabel.setWordWrap(True)
        self.temporaryLabel.setObjectName("temporaryLabel")
        self.errorsPlainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.errorsPlainTextEdit.setGeometry(QtCore.QRect(10, 300, 721, 191))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(12)
        self.errorsPlainTextEdit.setFont(font)
        self.errorsPlainTextEdit.setObjectName("errorsPlainTextEdit")
        self.styleNameLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.styleNameLineEdit.setGeometry(QtCore.QRect(10, 300, 500, 45))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        self.styleNameLineEdit.setFont(font)
        self.styleNameLineEdit.setObjectName("styleNameLineEdit")
        self.backButton = QtWidgets.QPushButton(self.centralwidget)
        self.backButton.setGeometry(QtCore.QRect(10, 30, 131, 51))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.backButton.setFont(font)
        self.backButton.setStyleSheet("background-color: rgb(200, 220, 210);")
        self.backButton.setDefault(False)
        self.backButton.setFlat(False)
        self.backButton.setObjectName("backButton")
        self.continueButton = QtWidgets.QPushButton(self.centralwidget)
        self.continueButton.setGeometry(QtCore.QRect(300, 530, 200, 50))
        font = QtGui.QFont()
        font.setFamily("Comic Sans MS")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.continueButton.setFont(font)
        self.continueButton.setStyleSheet("background-color: rgb(180, 197, 255);")
        self.continueButton.setDefault(False)
        self.continueButton.setFlat(False)
        self.continueButton.setObjectName("continueButton")
        createFileStyleWindow.setCentralWidget(self.centralwidget)

        self.temporaryLabel.hide()
        self.errorsPlainTextEdit.hide()
        self.styleNameLineEdit.hide()
        self.continueButton.hide()

        def inputTest():
            self.errorsPlainTextEdit.setPlainText('')
            global trueValueArr
            fileName, isTrueInput = inputFileStyleBlank()
            if isTrueInput:
                from WorkWithFileCreate.TestFillBlankFile import testFillBlank
                errArr, trueValueArr = testFillBlank(fileName)
                if errArr is None and trueValueArr is None:
                    editText("Ошибка! Выбранный файл поврежден или не является бланком форматирования стилей")
                else:
                    if not errArr:
                        self.temporaryLabel.setText("Введите имя стиля:")
                        self.temporaryLabel.show()
                        self.styleNameLineEdit.show()
                        self.continueButton.show()
                        self.errorsPlainTextEdit.hide()
                        self.inputFilesButton.hide()
                    else:
                        for err in errArr:
                            self.errorsPlainTextEdit.appendHtml(err)
                        self.temporaryLabel.setText("Список ошибок:")
                        self.temporaryLabel.show()
                        self.errorsPlainTextEdit.show()

        self.backButton.clicked.connect(backWindow)
        self.inputFilesButton.clicked.connect(inputTest)
        self.continueButton.clicked.connect(lambda: inputNameStileFormattingTest(self.styleNameLineEdit.text()))

        import webbrowser
        webbrowser.open('OutputFiles\\БланкФорматирования.docx')

        self.retranslateUi(createFileStyleWindow)
        QtCore.QMetaObject.connectSlotsByName(createFileStyleWindow)

    def retranslateUi(self, createFileStyleWindow):
        _translate = QtCore.QCoreApplication.translate
        createFileStyleWindow.setWindowTitle(_translate("createFileStyleWindow", "Experiment"))
        self.hiLabel.setText(_translate("createFileStyleWindow", "Создать стиль оформления"))
        self.infoLabel.setText(_translate("createFileStyleWindow",
                                          "Пожалуйста заполните форму. Затем введите имя стиля и, сохранив заполненный"
                                          " бланк,  загрузите, нажав соответсвующую кнопку"))
        self.inputFilesButton.setText(_translate("createFileStyleWindow", "Выбрать бланк"))
        self.temporaryLabel.setText(_translate("createFileStyleWindow", "Список ошибок:"))
        self.styleNameLineEdit.setPlaceholderText(_translate("createFileStyleWindow", "Пример: example file style"))
        self.backButton.setText(_translate("createFileStyleWindow", "Назад"))
        self.continueButton.setText(_translate("createFileStyleWindow", "Продолжить"))


def inputFileStyleBlank():
    isTrueInput = True
    fileName = QFileDialog.getOpenFileName(None, "Выберете файл", "", "All Files (*)")[0]
    if fileName != "":
        if fileName[fileName.rfind(".") + 1:] != "docx":
            isTrueInput = False
            editText(f'Ошибка! Выбранный формат файла "{fileName[fileName.find("."):]}" не поддерживается, необходим '
                     '"docx" файл')
    else:
        isTrueInput = False
        editText('Ошибка! Вы не выбрали ни одного файла для загрузки')
    return fileName, isTrueInput


def inputNameStileFormattingTest(styleName):
    x = re.search("[/|\\|:|*|?|\"|<|>|\\|]", styleName)
    if x:
        editText("Ошибка! В названии стиля нельзя использовать \\/:*?\"<>|")
    else:
        if styleName == '':
            editText("Ошибка! Название стиля не может быть пустым")
        else:
            if findIdPatternByName(styleName):
                editText("Ошибка! Стиль с таким именем уже существует")
            else:
                createNewPattern(styleName)
                fillCollectionInPattern(findIdPatternByName(styleName)[0][0])
                fillFormattingUnit(trueValueArr)
                editText("Стиль успешно создан!")
                from Windows.QtWindows.menuActionsWindow import showMenuActionsWindow
                showMenuActionsWindow()
                closeCreateFileStyleWindow()


def backWindow():
    from Windows.QtWindows.menuActionsWindow import showMenuActionsWindow
    showMenuActionsWindow()
    closeCreateFileStyleWindow()


def showCreateFileStyleWindow():
    global createFileStyleWindow
    createFileStyleWindow = QtWidgets.QMainWindow()
    ui = Ui_createFileStyleWindow()
    ui.setupUi(createFileStyleWindow)
    createFileStyleWindow.show()


def closeCreateFileStyleWindow():
    createFileStyleWindow.close()
