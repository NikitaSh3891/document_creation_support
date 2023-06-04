import os
import comtypes.client

from base64 import b64encode, b64decode
from io import BytesIO

"""
Класс предназначен для конвертации
"""


def convertFile(fileName, fileExtensionInput, fileExtensionOutput):
    try:
        # https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
        if fileExtensionOutput == ".doc":
            wdFormat = 0
        elif fileExtensionOutput == ".txt":
            wdFormat = 2
        elif fileExtensionOutput == ".dos":
            wdFormat = 4
        elif fileExtensionOutput == ".rtf":
            wdFormat = 6
        elif fileExtensionOutput == ".html":
            wdFormat = 8
        elif fileExtensionOutput == ".docx":
            wdFormat = 16
        elif fileExtensionOutput == ".pdf":
            wdFormat = 17
        elif fileExtensionOutput == ".xps":
            wdFormat = 18
        elif fileExtensionOutput == ".xml":
            wdFormat = 19
        elif fileExtensionOutput == ".odf":
            wdFormat = 23
        else:
            wdFormat = -1
        if wdFormat != -1:
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(os.path.abspath('OutputFiles\\' + fileName + fileExtensionInput))
            doc.SaveAs(os.path.abspath('OutputFiles\\' + fileName + fileExtensionOutput), FileFormat=wdFormat)
            doc.Close()
            word.Quit()
    except:
        from Windows.QtWindows.errWindow import editText
        editText("Вы открыли файл, который используется в ходе работы системы. Закройте все файлы, которые могут"
                 " относится к системе и попробуйте выполнить операцию еще раз. Если проблема не исчезнет - "
                 "перезагрузите компьютер или снимите задачу с Microsoft Word")



def convertFileToDOCX(fileName):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(str(fileName.replace("/", "\\")))
    doc.SaveAs(fileName[:fileName.rfind(".")] + ".docx", FileFormat=16)
    doc.Close()
    word.Quit()


def convertFileToRTF(fileName):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(str(fileName.replace("/", "\\")))
    doc.SaveAs(fileName[:fileName.rfind(".")] + ".rtf", FileFormat=6)
    doc.Close()
    word.Quit()


def deleteRTFFile(fileName):
    path = os.path.join(fileName[:fileName.rfind(".")] + ".rtf")
    os.remove(path)


def deleteDOCXFile(fileName):
    path = os.path.join(fileName[:fileName.rfind(".")] + ".docx")
    os.remove(path)


def deleteFile(fileName, fileExtensionInput):
    try:
        os.remove(os.path.abspath('OutputFiles\\' + fileName + fileExtensionInput))
    except:
        from Windows.QtWindows.errWindow import editText
        editText("Вы открыли файл, который используется в ходе работы системы. Закройте все файлы, которые могут"
                 " относится к системе и попробуйте выполнить операцию еще раз. Если проблема не исчезнет - "
                 "перезагрузите компьютер или снимите задачу с Microsoft Word")


def convertImageToString(nameImg):
    file = open(nameImg, "rb")
    stringImg = b64encode(file.read())
    return stringImg


def convertStringToImage(imgString):
    img = BytesIO(b64decode(imgString))
    return img


def convertFileToString(fileName):
    f = open(fileName, "r")
    rtfStr = f.read()
    f.close()
    return rtfStr


def convertStringToFile(rtfStr, outputFileName):
    try:
        docOutput = open('OutputFiles\\' + outputFileName + ".rtf", "w")
        docOutput.write(rtfStr)
        docOutput.close()
    except:
        from Windows.QtWindows.errWindow import editText
        editText("Вы открыли файл, который используется в ходе работы системы. Закройте все файлы, которые могут"
                 " относится к системе и попробуйте выполнить операцию еще раз. Если проблема не исчезнет - "
                 "перезагрузите компьютер или снимите задачу с Microsoft Word")
