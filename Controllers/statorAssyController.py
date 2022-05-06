import sqlite3

class statorAssy:

    __connection = None

    def __init__(self) -> None:
        self.__connection = sqlite3.connect('Stator_Slot.db')

    def select_column(self):
        cursor = self.__connection.cursor()
        record = cursor.execute("SELECT * From Stator_Slot")

        if not record == None:
            return [l[0] for l in record.description]

    def select_data(self, statorAssy):
        cursor = self.__connection.cursor()
        cursor.execute("SELECT * From Stator_Slot WHERE New_SAP = '"+ statorAssy +"'")
        record = cursor.fetchone()

        if not record == None:
            return list(record)
        else:
            return False