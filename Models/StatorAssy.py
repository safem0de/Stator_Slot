
class StatorAssyDetail:

    def __init__(self) -> None:
        pass

    __New_SAP = ""
    __StatorAssy = ""
    __StackNo = ""
    __StackSAP = ""
    __Slot_1 = ""
    __Slot_1_SAP = ""
    __Slot_2 = ""
    __Slot_2_SAP = ""

    ## Getter ##

    def getNewSAP(self):
        return self.__New_SAP

    def getStatorAssy(self):
        return self.__StatorAssy

    def getStackNo(self):
        return self.__StackNo

    def getStackSAP(self):
        return self.__StackSAP

    def getSlot_1(self):
        return self.__Slot_1

    def getSlot_1_SAP(self):
        return self.__Slot_1_SAP

    def getSlot_2(self):
        return self.__Slot_2

    def getSlot_2_SAP(self):
        return self.__Slot_2_SAP
    
    ## Setter ##
    
    def setNewSAP(self, param):
        self.__New_SAP = param

    def setStatorAssy(self, param):
        self.__StatorAssy = param

    def setStackNo(self, param):
        self.__StackNo = param

    def setStackSAP(self, param):
        self.__StackSAP = param

    def setSlot_1(self, param):
        self.__Slot_1 = param

    def setSlot_1_SAP(self, param):
        self.__Slot_1_SAP = param

    def setSlot_2(self, param):
        self.__Slot_2 = param

    def setSlot_2_SAP(self, param):
        self.__Slot_2_SAP = param