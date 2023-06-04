import psycopg2
import cryptocode

from datetime import datetime


"""
Класс предназначен для подключения и работы с базой данных
"""


def connectToDB():
    try:
        docInput = open("config.txt", "r")
        infoArr = docInput.read().split("\n")
        con = psycopg2.connect(
            database=cryptocode.decrypt(infoArr[3], infoArr[0]),
            user=cryptocode.decrypt(infoArr[1], infoArr[0]),
            password=cryptocode.decrypt(infoArr[2], infoArr[0]),
            host=cryptocode.decrypt(infoArr[4], infoArr[0]),
            port=cryptocode.decrypt(infoArr[5], infoArr[0])
        )
        # con = psycopg2.connect(
        #     database="document_creation_support",
        #     user="nik",
        #     password="12345678",
        #     host="192.168.56.50",
        #     port="5432"
        # )
        return con
    except:
        # from InputTest import setTextCriticalErrWindow
        # setTextCriticalErrWindow('Ошибка! Соединение с базой данных не установлено')
        pass


def findUser(phoneNumber, password):
    try:
        con = connectToDB()
        cur = con.cursor()
        cur.execute(
            "SELECT registered_user.id_registered_user, registered_user.name_registered_user FROM registered_user"
            " INNER JOIN password_history ON password_history.id_registered_user = registered_user.id_registered_user "
            "WHERE registered_user.phone_number = '" + phoneNumber + "' AND password_history.password = '" + password +
            "' AND date_edit_password = (SELECT MAX(date_edit_password) FROM password_history WHERE id_registered_user"
            f" = (SELECT id_registered_user FROM registered_user WHERE phone_number = '{phoneNumber}'))")
        result = cur.fetchall()
        con.close()
        for row in result:
            if row is not None:
                return row
            else:
                return None
    except:
        from InputTest import setTextCriticalErrWindow
        setTextCriticalErrWindow('Ошибка! Соединение с базой данных не установлено')


def findExperimentByUser(idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT experiment.id_experiment, name_experiment, start_date_experiment, stop_date_experiment "
                "FROM experiment WHERE id_access_to_info <= (SELECT MAX(id_access_to_info) FROM type_role "
                "INNER JOIN user_role ON user_role.id_type_role = type_role.id_type_role "
                "INNER JOIN registered_user ON registered_user.id_registered_user = user_role.id_registered_user "
                f"WHERE registered_user.id_registered_user = {str(idUser)}) UNION "
                "SELECT experiment.id_experiment, name_experiment, start_date_experiment, stop_date_experiment "
                "FROM experiment INNER JOIN experiment_available_user ON "
                "experiment_available_user.id_experiment = experiment.id_experiment "
                f"WHERE id_registered_user = {str(idUser)}")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findExperimentAndFolderByUser(idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(
        "SELECT experiment.id_experiment, lvl_folder, name_folder, name_experiment, start_date_experiment, "
        "stop_date_experiment "
        "FROM folder_available_user INNER JOIN folder ON folder.id_folder = folder_available_user.id_folder "
        "LEFT JOIN experiment_in_folder ON experiment_in_folder.id_folder = folder.id_folder "
        "LEFT JOIN experiment ON experiment.id_experiment = experiment_in_folder.id_experiment "
        "WHERE id_registered_user = " + str(idUser))
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findNameExperiment(idExperiment):
    con = connectToDB()
    cur = con.cursor()
    cur.execute('SELECT name_experiment FROM experiment WHERE id_experiment = ' + str(idExperiment))
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findDataExperiment(idExperiment):
    con = connectToDB()
    cur = con.cursor()
    cur.execute('SELECT data_experiment.result_experiment, type_data.name_type_data FROM experiment INNER JOIN '
                'data_in_experiment ON data_in_experiment.id_experiment = experiment.id_experiment INNER JOIN '
                'data_experiment ON data_experiment.id_data_experiment = data_in_experiment.id_data_experiment '
                'INNER JOIN type_data ON type_data.id_type_data = data_experiment.id_type_data WHERE '
                'experiment.id_experiment = ' + str(idExperiment))

    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findStyleCollectionFromPattern(idPattern):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT style_collection.name_style_collection FROM pattern INNER JOIN collection_in_pattern ON "
                "collection_in_pattern.id_pattern = pattern.id_pattern INNER JOIN style_collection ON "
                "style_collection.id_style_collection = collection_in_pattern.id_style_collection "
                "WHERE pattern.id_pattern = '" + str(idPattern) + "'")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findFileStyleCollection(nameCollection, idPattern):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT formatting_unit.formatting FROM style_collection "
                "INNER JOIN formatting_collection ON formatting_collection.id_style_collection = style_collection.id_style_collection "
                "INNER JOIN formatting_unit ON formatting_unit.id_formatting_unit = formatting_collection.id_formatting_unit "
                "INNER JOIN type_formatting_unit ON type_formatting_unit.id_type_formatting_unit = formatting_unit.id_type_formatting_unit "
                "INNER JOIN collection_in_pattern ON collection_in_pattern.id_style_collection = style_collection.id_style_collection "
                "INNER JOIN pattern ON pattern.id_pattern = collection_in_pattern.id_pattern "
                f"WHERE style_collection.name_style_collection = '{str(nameCollection)}' AND pattern.id_pattern = '{str(idPattern)}' "
                "ORDER BY type_formatting_unit.id_type_formatting_unit")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findFileExtension():
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT id_file_extension, name_file_extension FROM file_extension")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findPhoneNumber(phoneNumber):
    try:
        con = connectToDB()
        cur = con.cursor()
        cur.execute(
            "SELECT id_registered_user, name_registered_user FROM registered_user WHERE phone_number = '"
            + phoneNumber + "'")
        result = cur.fetchall()
        con.close()
        for row in result:
            if row is not None:
                return row
            else:
                return None
    except:
        from InputTest import setTextCriticalErrWindow
        setTextCriticalErrWindow('Ошибка! Соединение с базой данных не установлено')


def findUserRole(idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT name_type_role FROM type_role INNER JOIN user_role ON user_role.id_type_role = "
                f"type_role.id_type_role WHERE id_registered_user = {idUser}")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findParamSystem(idUser, typeParamSystem):
    from InputTest import userRoleTest
    idUserRole = userRoleTest(idUser)
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT param FROM param_system WHERE id_type_role = {idUserRole} "
                f"AND id_type_param_system = {typeParamSystem} AND id_registered_user = {idUser}")
    result = cur.fetchall()
    con.close()
    if str(result) == "[]":
        result = findParamSystemWithOutIdUser(idUserRole, typeParamSystem)
    if result is not None:
        return result
    else:
        return None


def findParamSystemWithOutIdUser(idUserRole, typeParamSystem):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT param FROM param_system WHERE id_type_role = {idUserRole} "
                f"AND id_type_param_system = {typeParamSystem}")
    result = cur.fetchall()
    con.close()
    if str(result) == "[]":
        result = findParamSystemWithOutIdTypeRole(typeParamSystem)
    if result is not None:
        return result
    else:
        return None


def findParamSystemWithOutIdTypeRole(typeParamSystem):
    try:
        con = connectToDB()
        cur = con.cursor()
        cur.execute("SELECT param FROM param_system WHERE id_type_role IS NULL "
                    f"AND id_type_param_system = {typeParamSystem}")
        result = cur.fetchall()
        con.close()
        if result is not None:
            return result
        else:
            return None
    except AttributeError:
        return -1


def repeatPasswordTest(idUser, param):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT date_edit_password, password FROM password_history WHERE id_registered_user = {idUser} "
                f"ORDER BY date_edit_password DESC LIMIT {param}")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def resetPassword(idUser, password):
    date = str(datetime.now())[:-7]
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"INSERT INTO password_history (id_registered_user, date_edit_password, password) VALUES ({idUser},"
                f" '{date}','{password}');")
    con.commit()
    con.close()


def findLastTimeResetPass(idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT date_edit_password FROM password_history WHERE id_registered_user = {idUser} ORDER BY "
                f"date_edit_password DESC LIMIT 1")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findFileFormattingPattern():
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT id_pattern, name_pattern FROM pattern")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findTitlePage():
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT id_title_page, name_document FROM title_page "
                "INNER JOIN document ON document.id_document = title_page.id_document")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findTitlePageById(idTitlePage):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT document_text FROM title_page INNER JOIN document ON document.id_document = "
                f"title_page.id_document WHERE id_title_page = '{str(idTitlePage)}'")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findUserDocument(idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT id_document, name_document, date_create_document FROM document WHERE id_registered_user"
                f" = {str(idUser)}")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findAvailableDocument(idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT id_document, name_document, name_registered_user, date_create_document FROM document "
                "INNER JOIN registered_user ON registered_user.id_registered_user = document.id_registered_user "
                "WHERE id_access_to_info <= (SELECT MAX(id_access_to_info) FROM type_role "
                "INNER JOIN user_role ON user_role.id_type_role = type_role.id_type_role "
                "INNER JOIN registered_user ON registered_user.id_registered_user = user_role.id_registered_user "
                f"WHERE registered_user.id_registered_user = {str(idUser)}) "
                f"AND document.id_registered_user <> {str(idUser)} UNION "
                "SELECT document.id_document, document.name_document, registered_user.name_registered_user, "
                "document.date_create_document FROM document "
                "INNER JOIN document_available_user ON document_available_user.id_document = document.id_document "
                "INNER JOIN registered_user ON registered_user.id_registered_user = document.id_registered_user "
                f"WHERE document_available_user.id_registered_user = {str(idUser)} "
                f"AND document.id_registered_user <> {str(idUser)}")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findNameAndTextDocument(idDocument):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT name_document, document_text FROM document WHERE id_document = '{str(idDocument)}'")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def fillDocumentToDB(idUser, nameDoc, idAccess, idFileExtension, document, idPattern):
    date = str(datetime.now())[:-7]
    con = connectToDB()
    cur = con.cursor()
    if idPattern is None:
        cur.execute("INSERT INTO document (name_document, date_create_document, id_access_to_info, id_file_extension, "
                    f"document_text, id_registered_user, id_pattern) VALUES "
                    f"('{nameDoc[nameDoc.rfind('/') + 1:nameDoc.rfind('.')]}', '{str(date)}', '{str(idAccess)}', "
                    f"'{str(idFileExtension)}', '{str(document)}', '{str(idUser)}', null)")
    else:
        cur.execute("INSERT INTO document (name_document, date_create_document, id_access_to_info, id_file_extension, "
                    f"document_text, id_registered_user, id_pattern) VALUES "
                    f"('{nameDoc[nameDoc.rfind('/') + 1:nameDoc.rfind('.')]}', '{str(date)}', '{str(idAccess)}', "
                    f"'{str(idFileExtension)}', '{str(document)}', '{str(idUser)}', '{str(idPattern)}')")
    con.commit()
    con.close()


def fillDocAvailableUser(idDoc, idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("INSERT INTO document_available_user (id_document, id_registered_user) VALUES "
                f"('{str(idDoc)}', '{str(idUser)}')")
    con.commit()
    con.close()


def fillExperimentInDocument(idDoc, idExperiment):
    con = connectToDB()
    cur = con.cursor()
    cur.execute("INSERT INTO experiment_in_document (id_document, id_experiment) VALUES "
                f"('{str(idDoc)}', '{str(idExperiment)}')")
    con.commit()
    con.close()


def findIDLastLoadDocument(idUser):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT id_document FROM document  WHERE id_registered_user = '{str(idUser)}'"
                "ORDER BY date_create_document DESC LIMIT 1")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findAccessToInfo():
    con = connectToDB()
    cur = con.cursor()
    cur.execute("SELECT id_access_to_info, name_access_to_info FROM access_to_info")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findPositionCellInBlank(idBlank):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT position_cell FROM cell_in_blank WHERE id_blank = '{str(idBlank)}'")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def findIdPatternByName(namePattern):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT id_pattern FROM pattern WHERE name_pattern = '{namePattern}'")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def createNewPattern(namePattern):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"INSERT INTO pattern (name_pattern) VALUES ('{str(namePattern)}')")
    con.commit()
    con.close()
    fillStyleCollection()


def fillStyleCollection():
    con = connectToDB()
    cur = con.cursor()
    cur.execute(
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Common'); \n"
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Normal'); \n"
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Heading 1'); \n"
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Heading 2'); \n"
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Heading 3'); \n"
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Caption'); \n"
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Table'); \n"
        f"INSERT INTO style_collection (name_style_collection) VALUES ('Image'); \n")
    con.commit()
    con.close()


def findLastIdStyleCollection():
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT id_style_collection FROM style_collection ORDER BY id_style_collection DESC LIMIT 1")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def fillCollectionInPattern(idPattern):
    firstId = findLastIdStyleCollection()[0][0]
    insertString = "INSERT INTO collection_in_pattern (id_pattern, id_style_collection) VALUES ("
    resultString = ""
    count = firstId - 7
    while count <= firstId:
        resultString += insertString + f"'{str(idPattern)}', '{str(count)}'); \n"
        count += 1
    con = connectToDB()
    cur = con.cursor()
    cur.execute(resultString)
    con.commit()
    con.close()


def fillFormattingUnit(formattingArr):
    insertString = "INSERT INTO formatting_unit (id_type_formatting_unit, formatting) VALUES ("
    resultString = ""
    count = 1
    for formatting in formattingArr:
        if count == 16:
            count = 5
            resultString += insertString + f"'{str(count)}', '{str(formatting)}'); \n"
        else:
            resultString += insertString + f"'{str(count)}', '{str(formatting)}'); \n"
        count += 1
    con = connectToDB()
    cur = con.cursor()
    cur.execute(resultString)
    con.commit()
    con.close()
    fillFormattingCollection()


def findLastIdFormattingUnit():
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT id_formatting_unit FROM formatting_unit ORDER BY id_formatting_unit DESC LIMIT 1")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def fillFormattingCollection():
    lastIdFormatting = findLastIdFormattingUnit()[0][0]
    insertString = "INSERT INTO formatting_collection (id_style_collection, id_formatting_unit) VALUES ("
    resultString = ""
    count = lastIdFormatting - 80
    idStyleCollection = findLastIdStyleCollection()[0][0]
    countIdStyleCollection = idStyleCollection - 7
    valueAddIdStyleCollection = 1
    while count <= lastIdFormatting:
        if valueAddIdStyleCollection == 1:
            for i in range(4):
                resultString += insertString + f"'{str(countIdStyleCollection)}', '{str(count)}'); \n"
                count += 1
        else:
            for i in range(11):
                resultString += insertString + f"'{str(countIdStyleCollection)}', '{str(count)}'); \n"
                count += 1
        valueAddIdStyleCollection += 1
        countIdStyleCollection += 1
    con = connectToDB()
    cur = con.cursor()
    cur.execute(resultString)
    con.commit()
    con.close()


def findLoggingPatter(typeLogging):
    con = connectToDB()
    cur = con.cursor()
    cur.execute(f"SELECT name_logging_pattern FROM logging_pattern WHERE id_type_logging = '{str(typeLogging)}'")
    result = cur.fetchall()
    con.close()
    if result is not None:
        return result
    else:
        return None


def loggingInfo(idTypeLogging, idUser, log, idLvlLogging, idLoggingPattern):
    date = str(datetime.now())[:-7]
    con = connectToDB()
    cur = con.cursor()
    cur.execute(
        f"INSERT INTO logging (id_type_logging, id_registered_user, date_logging, log, id_lvl_logging, "
        f"id_logging_pattern) VALUES ('{str(idTypeLogging)}', {str(idUser)}, '{str(date)}', '{str(log)}', "
        f"'{str(idLvlLogging)}', '{str(idLoggingPattern)}')")
    con.commit()
    con.close()
