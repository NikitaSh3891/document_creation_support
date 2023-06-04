import numpy as np
import matplotlib.pyplot as plt

from WorkWithFileCreate.CreateBlankFile import createEmptyBlank

"""
Класс предназначен для создания графиков
"""


def createStat(fileName):
    phrase = "[IdErr-"
    statArr = np.zeros(12)
    fName = fileName[fileName.rfind("/") + 1:fileName.rfind(".")]
    docOutput = open('OutputFiles/Результаты проверки файла/' + fName + ".txt")
    for line in docOutput:
        statArr[int(line[len(phrase): len(phrase) + 2].replace("]", "")) - 1] += 1
    docOutput.close()
    editStatImg(statArr, fName)
    createEmptyBlank(fName, statArr)


def editStatImg(statArr, fileName):
    paramArr = ["Шрифт",
                "Размера шрифта",
                "Межстрочного интервала",
                "Выравнивания текста",
                "Отступа первой строки",
                "Выделение текста",
                "Отступа послед.",
                "Отступа предыдущ.",
                "Разрыва абзаца",
                "Заглавные буквы",
                "Цвет текста",
                "Отступы полей "]

    labels = []
    resultStat = []
    for i in range(len(statArr)):
        if statArr[i] != 0:
            labels.append(paramArr[i])
            resultStat.append(int(statArr[i]))
    plt.pie(resultStat, labels=labels, autopct='%.0f%%')
    plt.savefig(f'OutputFiles\\Результаты проверки файла\\{fileName}.png')
