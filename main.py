import os

"""
Входная точка в программу
"""

if __name__ == '__main__':
    if not os.path.exists('OutputFiles'):
        os.mkdir("OutputFiles")
        os.mkdir("OutputFiles\\Результаты проверки файла")
    from Windows.QtWindows.mainWindow import start
    start()
