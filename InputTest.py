import re
import os.path
import hashlib
import webbrowser

import Converter
import WorkWithDB
import WorkWithFile
from WorkWithFileCreate.EditFileStyle import replaceSymbol

"""
Класс предназначен для проверок полеченных от пользователя параметров 
"""


def userRoleTest(idUser):
    userRoleArr = WorkWithDB.findUserRole(idUser)
    idUserRole = 0
    for userRole in userRoleArr:
        if idUserRole != 2 and userRole == "Пользователь":
            idUserRole = 1
        if userRole == "Администратор":
            idUserRole = 2
    return idUserRole


def haveExperimentByUserTest(idUser, nameUser):
    experiments = WorkWithDB.findExperimentByUser(idUser)
    folders = WorkWithDB.findExperimentAndFolderByUser(idUser)
    if len(experiments) > 0 or len(folders) > 0:
        from Windows.QtWindows.tableExperimentsAndFoldersWindow import editNameUser, editExperiments, \
            showTableExperimentWindow
        editNameUser("Здравствуйте, " + ''.join(nameUser))
        editExperiments(experiments, folders)
        showTableExperimentWindow()
        from Windows.QtWindows.menuActionsWindow import closeMenuActionsWindow
        closeMenuActionsWindow()
    else:
        setTextErrWindow('Ошибка! Вы не принимали участие в экспериментах!')


def userTest(phoneNumber, password):
    phoneNumber = re.sub("(\\s|\\(|\\)|-|\\+)", "", phoneNumber)
    resultPhoneNumberTest = phoneNumberTest(phoneNumber)
    if resultPhoneNumberTest != -1:
        if phoneNumberTest(phoneNumber):
            password = strPasswordToHash(password)
            user = WorkWithDB.findUser(phoneNumber, password)
            if user is not None:
                global idUser
                idUser = user[0]
                WorkWithDB.loggingInfo(1, idUser, "", 4, 1)
                WorkWithFile.editParamUser(user)
                from Windows.QtWindows.menuActionsWindow import showMenuActionsWindow, editValueUser
                editValueUser(user)
                showMenuActionsWindow()
                from Windows.QtWindows.mainWindow import closeMainWindow
                closeMainWindow()
            else:
                WorkWithDB.loggingInfo(5, "null", "", 4, 5)
                setTextErrWindow('Ошибка! Пользователя с такими данными не найдено')
        else:
            setTextErrWindow('Ошибка! Вы неверно ввели номер телефона')


def resetPasswordTest(phoneNumber, password, passwordTwo):
    phoneNumber = re.sub("(\\s|\\(|\\)|-|\\+)", "", phoneNumber)
    resultPhoneNumberTest = phoneNumberTest(phoneNumber)
    if resultPhoneNumberTest != -1:
        if phoneNumberTest(phoneNumber):
            if passwordTest(password):
                if password == passwordTwo:
                    user = WorkWithDB.findPhoneNumber(phoneNumber)
                    if user is not None:
                        idUser = user[0]
                        password = strPasswordToHash(password)
                        if timeResetPasswordTest(idUser):
                            countPassParam = replaceSymbol(WorkWithDB.findParamSystem(idUser, 9))
                            passwordHistory = WorkWithDB.repeatPasswordTest(idUser, countPassParam)
                            isTrue = False
                            for passHisLine in passwordHistory:
                                passHisLine = str(passHisLine)
                                if passHisLine[passHisLine.find("'")+1:-2] == password:
                                    isTrue = True
                            if isTrue:
                                setTextErrWindow('Ошибка! Вы уже недавно использовали этот пароль\n'
                                                 'Пожалуйста, придумайте новый!')
                            else:
                                WorkWithDB.resetPassword(idUser, password)
                                setTextErrWindow('Успешно! Ваш пароль изменен')
                                from Windows.QtWindows.resetPasswordWindow import closeResetPasswordWindow
                                closeResetPasswordWindow()
                                from Windows.QtWindows.mainWindow import showMainWindow
                                showMainWindow()
                        else:
                            setTextErrWindow('Ошибка! Вы уже изменяли пароль недавно')
                    else:
                        setTextErrWindow('Ошибка! Пользователя с такими данными не найдено')
                else:
                    setTextErrWindow('Ошибка! Пароли не совпадают!')
            else:
                setTextErrWindow('Ошибка! Пароль не соответствует маске!')
        else:
            setTextErrWindow('Ошибка! Вы неверно ввели номер телефона')


def strPasswordToHash(strPassword):
    hashOne = hashlib.md5()
    hashOne.update(strPassword.encode('utf-8'))
    password = hashOne.hexdigest()
    password = password[::-1]
    hashTwo = hashlib.md5()
    hashTwo.update(password.encode('utf-8'))
    password = hashTwo.hexdigest()
    return password


def setTextErrWindow(errText):
    from Windows.QtWindows.errWindow import editText
    editText(errText)


def setTextCriticalErrWindow(errText):
    from Windows.QtWindows.criticalErrWindow import editText
    editText(errText)
    from Windows.QtWindows.criticalErrWindow import showErrWindow
    showErrWindow()


def phoneNumberTest(phoneNumber):
    regexParam = WorkWithDB.findParamSystemWithOutIdTypeRole(5)
    if regexParam == -1:
        setTextErrWindow("Ошибка! Не удалось установить подключение с базой данных")
        return -1
    else:
        x = re.search(str(regexParam)[3:-4].replace("\\\\", "\\"), phoneNumber)
        return x


def passwordTest(password):
    regexParam = WorkWithDB.findParamSystemWithOutIdTypeRole(4)
    x = re.search(str(regexParam)[3:-4], password)
    return x


def timeResetPasswordTest(idUser):
    isTrue = False
    timeParam = replaceSymbol(WorkWithDB.findParamSystem(idUser, 1))
    lastTimeResetPass = replaceSymbol(WorkWithDB.findLastTimeResetPass(idUser))[17:]
    from datetime import datetime
    lastTimeResetPass = datetime.strptime(lastTimeResetPass, '%Y %m %d %H %M %S')
    differenceDate = str(datetime.now().date() - lastTimeResetPass.date())
    if re.search(str("[0-9]{1,} days"), str(differenceDate)):
        differenceDate = differenceDate[:differenceDate.find(" ")]
        if int(differenceDate) > int(timeParam):
            isTrue = True
    return isTrue


def fileNameTest(fileName):
    x = re.search("[/|\\|:|*|?|\"|<|>|\\|]", fileName)
    if x:
        setTextErrWindow("Ошибка! В имени файла нельзя использовать \\/:*?\"<>|")
    else:
        if fileName == '':
            setTextErrWindow("Ошибка! Имя файла не может быть пустым")
        else:
            if os.path.exists("OutputFiles/" + fileName + ".docx"):
                from Windows.QtWindows.okOrCancelWindow import editTextAndFileName
                editTextAndFileName("Файл с таким именем уже существует! Желаете заменить его?", fileName)
            else:
                tryToCreateDocument(fileName)


def replaceFileTest(isReplace, fileName):
    if isReplace:
        tryToCreateDocument(fileName)


def tryToCreateDocument(fileName):
    if isUploadFile:
        from Windows.QtWindows.errWindow import editText, closeErrWindow
        for idDocument in documentsCheckedList:
            documentArr = WorkWithDB.findNameAndTextDocument(idDocument)
            editText("Обрабатывается файл - " + documentArr[0][0])
            Converter.convertStringToFile(documentArr[0][1].replace('"', '\''), documentArr[0][0])
            isRTF = True
            for nameExtensions in extensionCheckedList:
                if nameExtensions != "rtf":
                    Converter.convertFile(documentArr[0][0], ".rtf", "." + nameExtensions)
                else:
                    isRTF = False
            if isRTF:
                Converter.deleteFile(documentArr[0][0], ".rtf")
            closeErrWindow()
        setTextErrWindow('Выгрузка произошла успешно!')
    else:
        if titlePage == -1:
            if WorkWithFile.createDocument(fileName, experimentsCheckedList, fileFormatting):
                setTextErrWindow('Файл успешно создан!')
                from Windows.QtWindows.inputParamFileCreateWindow import closeInputParamFileCreateWindow
                closeInputParamFileCreateWindow()
                from Windows.QtWindows.menuActionsWindow import showMenuActionsWindow
                showMenuActionsWindow()
            else:
                setTextErrWindow('Ошибка! Не удалось создать файл так как он используется другой программой')
        else:
            randomFileName = generateRandomString(10)
            Converter.convertStringToFile(WorkWithDB.findTitlePageById(titlePage)[0][0].replace('"', '\''), randomFileName)
            Converter.convertFile(randomFileName, ".rtf", ".docx")
            Converter.deleteFile(randomFileName, ".rtf")
            if WorkWithFile.createDocumentWithTitlePage(fileName, randomFileName + ".docx", experimentsCheckedList,  fileFormatting):
                Converter.deleteFile(randomFileName, ".docx")
                setTextErrWindow('Файл успешно создан!')
                from Windows.QtWindows.inputParamFileCreateWindow import closeInputParamFileCreateWindow
                closeInputParamFileCreateWindow()
                from Windows.QtWindows.menuActionsWindow import showMenuActionsWindow
                showMenuActionsWindow()
            else:
                setTextErrWindow('Ошибка! Не удалось создать файл так как он используется другой программой')
    path = os.path.abspath('OutputFiles')
    webbrowser.open(path)


def tryToLoadFile():
    from PyQt5.QtWidgets import QFileDialog
    fileNames = QFileDialog.getOpenFileNames(None, "Выберете файл", "", "All Files (*)")[0]
    if fileNames is not None:
        fileExtensions = WorkWithDB.findFileExtension()
        for fName in fileNames:
            isTrue = True
            for extension in fileExtensions:
                if fName[fName.rfind(".") + 1:] == extension[1]:
                    isTrue = False
                    break
            if isTrue:
                setTextErrWindow(f'Ошибка! Выбранный формат файла "{fName[fName.find("."):]}" не поддерживается')
            else:
                from Windows.QtWindows.inputFileAccessWindows import showInputFileAccessWindows, editParamFileLoad
                editParamFileLoad(idUser, fileNames)
                showInputFileAccessWindows()
                from Windows.QtWindows.menuActionsWindow import closeMenuActionsWindow
                closeMenuActionsWindow()
    else:
        setTextErrWindow('Ошибка! Вы не выбрали ни одного файла для загрузки')


def tryToFormattingFile(checkTestArr):
    from WorkWithFileCreate.CheckFileForFormatting import editNumCharacter, startCheckForFileFormatting
    fileNames = fileNameFormatting.split("\n")
    editNumCharacter(10)
    for fName in fileNames:
        from datetime import datetime
        start_time = datetime.now()
        startCheckForFileFormatting(fName, checkTestArr, idPatternFormatting)
        print(datetime.now() - start_time)
    from Windows.QtWindows.menuActionsWindow import showMenuActionsWindow
    showMenuActionsWindow()
    if isEditFileFormatting:
        from WorkWithFileCreate.CheckAndEditFileForFormatting import editNumCharacter, \
            startCheckAndEditFileForFormatting
        fileNames = fileNameFormatting.split("\n")
        editNumCharacter(10)
        for fName in fileNames:
            startCheckAndEditFileForFormatting(fName, checkTestArr, idPatternFormatting)
        from Windows.QtWindows.menuActionsWindow import showMenuActionsWindow
        showMenuActionsWindow()
    if isCreateStatByFormatting:
        from WorkWithFileCreate.StatByFormatFileTest import createStat
        fileNames = fileNameFormatting.split("\n")
        editNumCharacter(10)
        for fName in fileNames:
            createStat(fName)

    setTextErrWindow("Файлы были успешно проверены!")
    path = os.path.abspath('OutputFiles/Результаты проверки файла')
    webbrowser.open(path)


def generateRandomString(length):
    import random
    import string

    letters = string.ascii_lowercase
    randomString = ''.join(random.choice(letters) for i in range(length))
    return randomString


def getExperimentsCheckedList(checkedList):
    global experimentsCheckedList
    experimentsCheckedList = checkedList


def getDocumentsCheckedList(checkedList):
    global documentsCheckedList
    documentsCheckedList = checkedList


def editExtensionCheckedList(checkedList):
    global extensionCheckedList
    extensionCheckedList = checkedList


def editIsUploadFile(value):
    global isUploadFile
    isUploadFile = value


def editFileFormatting(value):
    global fileFormatting
    fileFormatting = value


def editTitlePage(value):
    global titlePage
    titlePage = value


def editTypeFileFormatting(_isEditFileFormatting, _isCreateStatByFormatting, _fileNameFormatting, _idPatternFormatting):
    global isEditFileFormatting
    isEditFileFormatting = _isEditFileFormatting
    global isCreateStatByFormatting
    isCreateStatByFormatting = _isCreateStatByFormatting
    global fileNameFormatting
    fileNameFormatting = _fileNameFormatting
    global idPatternFormatting
    idPatternFormatting = _idPatternFormatting
