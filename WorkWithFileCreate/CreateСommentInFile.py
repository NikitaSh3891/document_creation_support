from thefuzz import fuzz
import ctypes

"""
Класс предназначен для создания комментариев в файле
"""


# Определение имени автора комментариев
def findDisplayName():
    userNameEx = ctypes.windll.secur32.GetUserNameExW
    nameDisplay = 3
    size = ctypes.pointer(ctypes.c_ulong(0))
    userNameEx(nameDisplay, None, size)
    name_buffer = ctypes.create_unicode_buffer(size.contents.value)
    userNameEx(nameDisplay, name_buffer, size)
    return name_buffer.value


# Добавление комментария
def addComment(msgText, paragraphText, paragraphs, nameUser):
    for paragraph in paragraphs:
        if type(paragraph) == list:
            tableText = [p.text for p in paragraph]
            tableText = ''.join(tableText)
            if len(tableText) >= len(paragraphText)-5:
                res = fuzz.partial_ratio(paragraphText.lower(), tableText.lower())
                if res >= 97:
                    p = paragraph[-1]
                    run = p.add_run()
                    run.add_comment(msgText, author=nameUser)
        else:
            if len(paragraph.text) >= len(paragraphText):
                res = fuzz.partial_ratio(paragraphText.lower(), paragraph.text.lower())
                if res >= 97:
                    paragraph.add_comment(msgText, author=nameUser)
