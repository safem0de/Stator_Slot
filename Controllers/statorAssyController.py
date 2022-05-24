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

    def insertVaribleIntoTable(self, New_SAP, Statorassy, StackNo, StackSAP, Slot_1, Slot_1_SAP, Slot_2, Slot_2_SAP):
        try:
            cursor = self.__connection.cursor()
            print("Connected to SQLite")

            sqlite_insert_with_param = """INSERT INTO Stator_Slot
                            (New_SAP, Statorassy, StackNo, StackSAP, Slot_1, Slot_1_SAP, Slot_2, Slot_2_SAP) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?);"""

            data_tuple = (New_SAP, Statorassy, StackNo, StackSAP, Slot_1, Slot_1_SAP, Slot_2, Slot_2_SAP)
            cursor.execute(sqlite_insert_with_param, data_tuple)
            self.__connection.commit()
            print("Python Variables inserted successfully into SqliteDb_developers table")

            cursor.close()

        except sqlite3.Error as error:
            print("Failed to insert Python variable into sqlite table", error)
        finally:
            if self.__connection:
                self.__connection.close()
                print("The SQLite connection is closed")