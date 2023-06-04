import docx
import re

from WorkWithDB import findPositionCellInBlank

"""
Класс предназначен проверки полей бланк файла
"""


def testFillBlank(fileNameBlank):
    try:
        blankFile = open(fileNameBlank, 'rb')
        doc = docx.Document(blankFile)
        positionArr = findPositionCellInBlank(3)
        errArr = []
        trueValueArr = []
        for table in doc.tables:
            count = 0
            if table.cell(0, 1).text == "Common – Значение полей документа":
                while count < 4:
                    elemPos = str(positionArr[count][0]).split(" ")
                    if not re.search(elemPos[2].replace("\\\\", "\\"), table.cell(int(elemPos[0]), int(elemPos[1])).text):
                        errArr.append(f"Ошибка в таблице '{table.cell(0, 1).text}', в поле "
                                      f"'{table.cell(int(elemPos[0]), 0).text}'")
                    else:
                        trueValueArr.append(table.cell(int(elemPos[0]), int(elemPos[1])).text)
                    count += 1
            else:
                count = 4
                while count < len(positionArr):
                    elemPos = str(positionArr[count][0]).split(" ")
                    if not re.search(elemPos[2].replace("\\\\", "\\"), table.cell(int(elemPos[0]), int(elemPos[1])).text):
                        errArr.append(
                            f"Ошибка в таблице '{table.cell(0, 1).text}', в поле '{table.cell(int(elemPos[0]), 0).text}'")
                    else:
                        trueValueArr.append(table.cell(int(elemPos[0]), int(elemPos[1])).text)
                    count += 1
        blankFile.close()
        return errArr, trueValueArr
    except IndexError:
        return None, None
