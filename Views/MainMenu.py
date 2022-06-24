from datetime import datetime
from msilib.schema import Icon
from tkinter import font, ttk
from tkinter import *
import tkinter as tk

from tkinter import messagebox


# from ttkthemes import ThemedStyle
from Controllers.statorAssyController import statorAssy
from Models.StatorAssy import StatorAssyDetail

from pygame import mixer
from threading import Thread
import openpyxl
import pandas as pd
import re
import datetime
import os
import importlib

class MainMenu(ttk.Frame):

    mdl = StatorAssyDetail()
    ctrl = statorAssy()
    w = 25
    
    def __init__(self, parent):
        super().__init__(parent)

        if '_PYIBoot_SPLASH' in os.environ and importlib.util.find_spec("pyi_splash"):
            import PyInstaller as pyi_splash
            pyi_splash.update_text('UI Loaded ...')
            pyi_splash.close()
            # log.info('Splash screen closed.')

        # https://stackoverflow.com/questions/51697858/python-how-do-i-add-a-theme-from-ttkthemes-package-to-a-guizero-application
        # self.style = ThemedStyle(self)
        # self.style.set_theme("aquativo")

        #create widgets
        self.labelheader = ttk.Label(self, text = 'Insert Slot Process', font=("Comic Sans MS", 20))
        self.labelheader.grid(row=0, column=0, sticky=tk.W)

        self.lf = LabelFrame(self, text="Dashboard ",font=("Comic Sans MS", 12))
        self.lf.grid(row=1, column=0, columnspan=20, sticky=tk.W)

        self.alignments = ('Set Up', 'Add Data', 'Edit Data', 'Slot Output', 'Credits')
        self.nb = ttk.Notebook(self.lf)
        self.nb.grid(column=0, row=0, ipadx=10, ipady=10)

        self.f0 = Frame(self.nb, width=1024, height=280, name=self.alignments[0].replace(" ","_").lower())
        self.f1 = Frame(self.nb, width=1024, height=280, name=self.alignments[1].replace(" ","_").lower())
        self.f2 = Frame(self.nb, width=1024, height=280, name=self.alignments[2].replace(" ","_").lower())
        self.f3 = Frame(self.nb, width=1024, height=280, name=self.alignments[3].replace(" ","_").lower())
        self.f4 = Frame(self.nb, width=1024, height=280, name=self.alignments[4].replace(" ","_").lower())

        self.nb.add(self.f0, text=self.alignments[0])
        self.nb.add(self.f1, text=self.alignments[1])
        self.nb.add(self.f2, text=self.alignments[2])
        self.nb.add(self.f3, text=self.alignments[3])
        self.nb.add(self.f4, text=self.alignments[4])

        self.style = ttk.Style()
        self.style.configure('TNotebook.Tab', font=("Comic Sans MS", 14))

        ##################### ==== Credits ==== ######################
        # Update programing details and                              #
        # infomation of project maker                                #
        ############################################################## 

        self.f = open("credits.txt", encoding="utf8")
        # print(self.f.read())
        self.lblCredits = ttk.Label(self.f4, text=self.f.read())
        self.lblCredits.grid(row=0, column=0, padx=3, pady=3, sticky=tk.NE)

        #### ==== Table ==== ####

        self.LblTable = ttk.Label(self.f0, text = 'Table : ', font=("Comic Sans MS", 14))
        self.LblTable.grid(row=0, column=0, padx=3, pady=3, sticky=tk.NE)

        self.selected_Table = tk.StringVar()
        self.table_cb = ttk.Combobox(self.f0, textvariable=self.selected_Table, font=("Comic Sans MS", 14))
        self.table_cb.bind('<<ComboboxSelected>>', lambda event : onSelected_Table(event, self.selected_Table.get()))
        self.table_cb.bind('<Return>', lambda event : onSelected_Table(event, self.selected_Table.get()))
        self.table_cb.bind('<Tab>', lambda event : onSelected_Table(event, self.selected_Table.get()))
        self.table_cb.bind("<Button-1>", lambda event: focus_in(event, self.table_cb,self.canvasTable))
        self.table_cb['values'] = [f"TABLE{m}" for m in range(1, 21)]
        self.table_cb.grid(row=0, column=1, padx=3, pady=3, sticky=tk.NW)

        self.canvasTable = tk.Canvas(self.f0, width=self.w, height=self.w)
        self.canvasTable.grid(row=0, column=2, sticky=tk.EW)
        self.table_cb.focus()

        #### ==== Arranger ==== ####

        self.LblArranger = ttk.Label(self.f0, text = 'Arranger : ', font=("Comic Sans MS", 14))
        self.LblArranger.grid(row=1, column=0, padx=3, pady=3, sticky=tk.NE)

        self.Arranger = tk.StringVar()
        self.txtArranger = ttk.Entry(self.f0, textvariable=self.Arranger, font=("Comic Sans MS", 14))
        self.txtArranger.bind("<Return>", lambda event: onKeyPress_Arranger(event, self.Arranger.get()))
        self.txtArranger.bind("<Tab>", lambda event: onKeyPress_Arranger(event, self.Arranger.get()))
        self.txtArranger.bind("<Button-1>", lambda event: focus_in(event, self.txtArranger, self.canvasArranger))
        self.txtArranger.grid(row=1, column=1, padx=3, pady=3, sticky=tk.EW)

        self.canvasArranger = tk.Canvas(self.f0, width=self.w, height=self.w)
        self.canvasArranger.grid(row=1, column=2, sticky=tk.EW)

        #### ==== Operator ==== ####

        self.LblOperator = ttk.Label(self.f0, text = 'Operator : ', font=("Comic Sans MS", 14))
        self.LblOperator.grid(row=2, column=0, padx=3, pady=3, sticky=tk.NE)

        self.Operator = tk.StringVar()
        self.txtOperator = ttk.Entry(self.f0, textvariable=self.Operator, font=("Comic Sans MS", 14))
        self.txtOperator.bind("<Return>", lambda event: onKeyPress_Operator(event,self.Operator.get()))
        self.txtOperator.bind("<Tab>", lambda event: onKeyPress_Operator(event,self.Operator.get()))
        self.txtOperator.bind("<Button-1>", lambda event: focus_in(event,self.txtOperator,self.canvasOperator))
        self.txtOperator.grid(row=2, column=1, padx=3, pady=3, sticky=tk.EW)

        self.canvasOperator = tk.Canvas(self.f0, width=self.w, height=self.w)
        self.canvasOperator.grid(row=2, column=2, sticky=tk.EW)

        #### ==== Stator Assy ==== ####

        self.LblStatorAssy = ttk.Label(self.f0, text = 'Stator Assy : ', font=("Comic Sans MS", 14))
        self.LblStatorAssy.grid(row=0, column=3, padx=3, pady=3, sticky=tk.NW)
            
        self.statorAssy = tk.StringVar()
        self.txtStatorAssy = ttk.Entry(self.f0, textvariable=self.statorAssy, font=("Comic Sans MS", 14), name='sa')
        self.txtStatorAssy.bind("<Return>", lambda event: onclick_Loadmaster(event,self.statorAssy.get()))
        self.txtStatorAssy.bind("<Tab>", lambda event: onclick_Loadmaster(event,self.statorAssy.get()))
        self.txtStatorAssy.bind("<Button-1>", lambda event: focus_in(event,self.txtStatorAssy,self.canvasStatorAssy))
        self.txtStatorAssy.grid(row=0, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasStatorAssy = tk.Canvas(self.f0, width=self.w, height=self.w)
        self.canvasStatorAssy.grid(row=0, column=5, sticky=tk.EW)

        #### ==== Master ==== ####
        def MasterState():
            self.lf_master = LabelFrame(self.f0, text="Master",font=("Comic Sans MS", 14))
            self.lf_master.grid(row=4, column=0, columnspan=2, sticky=tk.N+tk.EW)

            self.LblStator_master = ttk.Label(self.lf_master, text = 'Stator :\t', font=("Comic Sans MS", 14))
            self.LblStator_master.grid(row=0, column=0, padx=3, pady=3, sticky=tk.NW)

            self.LblStator_Slot1 = ttk.Label(self.lf_master, text = 'Slot1 :\t', font=("Comic Sans MS", 14))
            self.LblStator_Slot1.grid(row=1, column=0, padx=3, pady=3, sticky=tk.NW)

            self.LblStator_Slot2 = ttk.Label(self.lf_master, text = 'Slot2 :\t', font=("Comic Sans MS", 14))
            self.LblStator_Slot2.grid(row=2, column=0, padx=3, pady=3, sticky=tk.NW)

            self.LblStator_master_SAP = ttk.Label(self.lf_master, text = 'SAP Stator : ', font=("Comic Sans MS", 14))
            self.LblStator_master_SAP.grid(row=3, column=0, padx=3, pady=3, sticky=tk.W)

            self.LblStator_Slot1_SAP = ttk.Label(self.lf_master, text = 'SAP Slot1 : ', font=("Comic Sans MS", 14))
            self.LblStator_Slot1_SAP.grid(row=4, column=0, padx=3, pady=3, sticky=tk.W)

            self.LblStator_Slot2_SAP = ttk.Label(self.lf_master, text = 'SAP Slot2 : ', font=("Comic Sans MS", 14))
            self.LblStator_Slot2_SAP.grid(row=5, column=0, padx=3, pady=3, sticky=tk.W)

        MasterState()

        #### ==== Slot1 ==== ####

        self.LblSlot1 = ttk.Label(self.f0, text = 'Slot1 : ', font=("Comic Sans MS", 14))
        self.LblSlot1.grid(row=1, column=3, padx=3, pady=3, sticky=tk.NE)

        self.slot1 = tk.StringVar()
        self.txtSlot1 = ttk.Entry(self.f0, textvariable=self.slot1, font=("Comic Sans MS", 14))
        self.txtSlot1.bind("<Return>", lambda event: onclick_slot1(event,self.slot1.get()))
        self.txtSlot1.bind("<Tab>", lambda event: onclick_slot1(event,self.slot1.get()))
        self.txtSlot1.bind("<Button-1>", lambda event: focus_in(event,self.txtSlot1,self.canvasSlot1))
        self.txtSlot1.grid(row=1, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasSlot1 = tk.Canvas(self.f0, width=self.w, height=self.w)
        self.canvasSlot1.grid(row=1, column=5, sticky=tk.EW)

        #### ==== Slot2 ==== ####

        self.LblSlot2 = ttk.Label(self.f0, text = 'Slot2 : ', font=("Comic Sans MS", 14))
        self.LblSlot2.grid(row=2, column=3, padx=3, pady=3, sticky=tk.NE)

        self.slot2 = tk.StringVar()
        self.txtSlot2 = ttk.Entry(self.f0, textvariable=self.slot2, font=("Comic Sans MS", 14))
        self.txtSlot2.bind("<Return>", lambda event: onclick_slot2(event,self.slot2.get()))
        self.txtSlot2.bind("<Tab>", lambda event: onclick_slot2(event,self.slot2.get()))
        self.txtSlot2.bind("<Button-1>", lambda event: focus_in(event,self.txtSlot2,self.canvasSlot2))
        self.txtSlot2.grid(row=2, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasSlot2 = tk.Canvas(self.f0, width=self.w, height=self.w)
        self.canvasSlot2.grid(row=2, column=5, sticky=tk.EW)

        #### ==== Stator Check ==== ####

        self.LblStator = ttk.Label(self.f0, text = 'Stator : ', font=("Comic Sans MS", 14))
        self.LblStator.grid(row=3, column=3, padx=3, pady=3, sticky=tk.NE)

        self.stator = tk.StringVar()
        self.txtStator = ttk.Entry(self.f0, textvariable=self.stator, font=("Comic Sans MS", 14))
        self.txtStator.bind("<Return>", lambda event: onclick_AddStator(event,self.stator.get()))
        self.txtStator.bind("<Tab>", lambda event: onclick_AddStator(event,self.stator.get()))
        self.txtStator.bind("<Button-1>", lambda event: focus_in(event,self.txtStator,self.canvasStator))
        self.txtStator.grid(row=3, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasStator = tk.Canvas(self.f0, width=self.w, height=self.w)
        self.canvasStator.grid(row=3, column=5, sticky=tk.EW)

        columns = ('Tray No.', 'Stator Stack')
        self.tree = ttk.Treeview(self.f0, columns=columns, show='headings')

        # define headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.CENTER)

        def item_selected(event):
            for selected_item in self.tree.selection():
                item = self.tree.item(selected_item)
                record = item['values']
                # show a message
                # showinfo(title='Information', message=record)
                MsgBox = messagebox.askquestion ('Delete Data',f'Are you sure to Delete {record[1]}',icon='warning')
                if MsgBox == 'yes':
                    x = self.tree.selection()
                    self.tree.delete(x)
                    messagebox.showinfo(title='Information', message=f'Tray {record[0]} : {record[1]} was deleted')
                else:
                    # messagebox.showinfo('Return','Not Delete')
                    pass

        self.tree.bind('<<TreeviewSelect>>', item_selected)

        self.tree.grid(row=4, column=4, padx=3, pady=3, sticky=tk.NSEW)

        ## add a scrollbar
        self.scrollbar = ttk.Scrollbar(self.f0, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=self.scrollbar.set)
        self.scrollbar.grid(row=4, column=5, padx=3, pady=3, sticky=tk.W + tk.NS)

        self.ClearBtn = ttk.Button(self.f0, text='ยืนยันการแก้ไข\nข้อผิดพลาด', command=lambda:unlock())
        self.ClearBtn.config(state="disabled")
        self.ClearBtn.grid(row=4, column=6, padx=3, pady=3, sticky=tk.S + tk.EW)

        def focus_in(event, Txt : ttk.Entry or ttk.Combobox, Canvas : tk.Canvas):
            
            if Txt.winfo_name() == 'sa': #ref(txtStatorAssy line:118)
                MasterState()

            Txt.delete(0,END)
            Canvas.delete('all')
            Txt.focus()

        def createGreenLight(cvs : tk.Canvas):
            x0, y0, x1, y1 = 2, 2, 22, 22
            cvs.create_oval(x0, y0, x1, y1, fill='green2')

        def onSelected_Table(event,value):
            try :
                match = re.search(r'\d{1,}', value)
                num = int(match.group())

                if "TABLE" in value and isinstance(num, int):
                    createGreenLight(self.canvasTable)
                    self.txtArranger.focus()
                else:
                    Thread(correct()).start()
                    Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                    Thread(focus_in(event,self.table_cb,self.canvasTable)).start()
                    return

            except Exception as e:
                Thread(correct()).start()
                Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                Thread(focus_in(event,self.table_cb,self.canvasTable)).start()
                return

        def onKeyPress_Arranger(event, value):
            
            x = re.findall("[A-Z]", value)
            y = re.findall('[0-9]+', value)

            if len(value) == 5 and len(x) and len(y):
                createGreenLight(self.canvasArranger)
                self.txtOperator.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                Thread(focus_in(event,self.txtArranger,self.canvasArranger)).start()
                return

        def onKeyPress_Operator(event, value):

            x = re.findall("[A-Z]", value)
            y = re.findall('[0-9]+', value)

            if len(value) == 5 and len(x) and len(y):
                createGreenLight(self.canvasOperator)
                self.txtStatorAssy.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                Thread(focus_in(event,self.txtOperator,self.canvasOperator)).start()
                return

        ### https://stackoverflow.com/questions/17125842/changing-the-text-on-a-label
        def onclick_Loadmaster(event, value):
            col_x = self.ctrl.select_column()
            x = self.ctrl.select_data(value)

            if x and len(value):
                df = pd.DataFrame(x,col_x).T
            
                self.mdl.setNewSAP(str(df['New_SAP'].values[0]))
                self.mdl.setStatorAssy(str(df['Statorassy'].values[0]))
                self.mdl.setStackNo(str(df['StackNo'].values[0]))
                self.mdl.setStackSAP(str(df['StackSAP'].values[0]))
                self.mdl.setSlot_1(str(df['Slot_1'].values[0]))
                self.mdl.setSlot_1_SAP(str(df['Slot_1_SAP'].values[0]))
                self.mdl.setSlot_2(str(df['Slot_2'].values[0]))
                self.mdl.setSlot_2_SAP(str(df['Slot_2_SAP'].values[0]))

                self.LblStator_master['text'] = f'Stator :\t{self.mdl.getStackNo()}'
                self.LblStator_Slot1['text'] = f'Slot1 :\t{self.mdl.getSlot_1()}'
                self.LblStator_Slot2['text'] = f'Slot2 :\t{self.mdl.getSlot_2()}'
                self.LblStator_master_SAP['text'] = f'SAP Stator : {self.mdl.getStackSAP()}'
                self.LblStator_Slot1_SAP['text'] = f'SAP Slot1 : {self.mdl.getSlot_1_SAP()}'
                self.LblStator_Slot2_SAP['text'] = f'SAP Slot2 : {self.mdl.getSlot_2_SAP()}'

                createGreenLight(self.canvasStatorAssy)
                self.txtSlot1.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Not Found!!! (ไม่พบข้อมูล)','กรุณาตรวจสอบ หรือ อาจเป็น Model.ใหม่?'],value)).start()
                Thread(focus_in(event, self.txtStatorAssy,self.canvasStatorAssy)).start()
                return
                

        def onclick_slot1(event, value):

            if value == self.mdl.getSlot_1_SAP() or value == self.mdl.getSlot_1():
                createGreenLight(self.canvasSlot1)
                self.txtSlot2.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Not Found!!! (ไม่พบข้อมูล)','กรุณาตรวจสอบ หรือ อาจเป็น Model.ใหม่?'],value)).start()
                Thread(focus_in(event, self.txtSlot1,self.canvasSlot1)).start()
                Thread(lockdown()).start()
                return

        def onclick_slot2(event, value):
            
            if value == self.mdl.getSlot_2_SAP() or value == self.mdl.getSlot_2():
                createGreenLight(self.canvasSlot2)
                self.txtStator.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Not Found!!! (ไม่พบข้อมูล)','กรุณาตรวจสอบ หรือ อาจเป็น Model.ใหม่?'],value)).start()
                Thread(focus_in(event, self.txtSlot2,self.canvasSlot2)).start()
                Thread(lockdown()).start()
                return

        def onclick_AddStator(event, value):

            if value == self.mdl.getStackSAP() or value == self.mdl.getStackNo():
                self.tree.insert('', tk.END, values=(len(self.tree.get_children())+1, value))
                self.txtStator.delete(0,END)
                self.txtStator.focus()

            elif checkEmpty() and (value == self.mdl.getNewSAP() or value == self.mdl.getStatorAssy()):
                # https://stackoverflow.com/questions/17977540/pandas-looking-up-the-list-of-sheets-in-an-excel-file
                createGreenLight(self.canvasStator)
                d = datetime.datetime.now()
                dt = d.strftime('%d-%m-%Y')
                dm = d.strftime('%b_%Y')

                excelName = f'SlotDailyRecord_{dm}.xlsx'

                x = None
                if self.mdl.getStackSAP() == '-':
                    x = self.mdl.getStackNo()
                else:
                    x = self.mdl.getStackSAP()

                if os.path.exists(excelName):
                    xl = pd.ExcelFile(excelName)
                    ls = xl.sheet_names

                    df = pd.DataFrame({
                            'Datetime': [d.strftime('%d-%m-%Y %H:%M')],
                            'Table': [self.selected_Table.get()],
                            'Arranger' : [self.Arranger.get()],
                            'Operator' : [self.Operator.get()],
                            'Stator Assy' : [self.statorAssy.get()],
                            'Slot 1' : [self.slot1.get()],
                            'Slot 2' : [self.slot2.get()],
                            'Stator' : [x],
                            'Tray Qty' : [len(self.tree.get_children())]
                            })
                    
                    if dt in ls:
                        # print('sheet exist')
                        # appending the data of df after the data of demo1.xlsx
                        with pd.ExcelWriter(excelName,mode="a",engine="openpyxl",if_sheet_exists="overlay") as writer:
                            df.to_excel(writer, sheet_name=dt,header=None, startrow=writer.sheets[dt].max_row, index=False)

                    else:
                        # print('sheet not exist')
                        # https://www.codegrepper.com/code-examples/python/pandas+to+excel+append+to+existing+sheet
                        wb = openpyxl.load_workbook(excelName)
                        writer = pd.ExcelWriter(excelName, engine = 'openpyxl')
                        writer.book = wb
                        df.to_excel(writer, sheet_name = dt, index=False)
                        writer.save()

                        # xl = pd.ExcelFile(excelName)
                        # ls = xl.sheet_names
                        # print(ls)
                else :
                    # print('create')
                    df = pd.DataFrame({
                        'Datetime': [d.strftime('%d-%m-%Y %H:%M')],
                        'Table': [self.selected_Table.get()],
                        'Arranger' : [self.Arranger.get()],
                        'Operator' : [self.Operator.get()],
                        'Stator Assy' : [self.statorAssy.get()],
                        'Slot 1' : [self.slot1.get()],
                        'Slot 2' : [self.slot2.get()],
                        'Stator' : [x],
                        'Tray Qty' : [len(self.tree.get_children())]
                   })

                    # Create a Pandas Excel writer using XlsxWriter as the engine.
                    writer = pd.ExcelWriter(excelName, engine='openpyxl')

                    # Convert the dataframe to an XlsxWriter Excel object.
                    df.to_excel(writer, sheet_name=dt, index=False)

                    # Close the Pandas Excel writer and output the Excel file.
                    writer.save()

                ClearAll()
                self.table_cb.focus()

            else:
                Thread(correct()).start()
                Thread(message_box(['Not Found!!! (ไม่พบข้อมูล)','กรุณาตรวจสอบ หรือ อาจเป็น Model.ใหม่?'],value)).start()
                Thread(focus_in(event, self.txtStator,self.canvasStator)).start()
                Thread(lockdown()).start()
                return

        def ClearAll():
            self.table_cb.delete(0,END)
            self.txtArranger.delete(0,END)
            self.txtOperator.delete(0,END)
            self.txtStatorAssy.delete(0,END)
            self.txtSlot1.delete(0,END)
            self.txtSlot2.delete(0,END)
            self.txtStator.delete(0,END)
            self.canvasTable.delete('all')
            self.canvasArranger.delete('all')
            self.canvasOperator.delete('all')
            self.canvasStatorAssy.delete('all')
            self.canvasSlot1.delete('all')
            self.canvasSlot2.delete('all')
            self.canvasStator.delete('all')
            MasterState()
            clearTreeview(self.tree)

        def clearTreeview(tree:ttk.Treeview):
            for i in tree.get_children():
                tree.delete(i)

        def correct():
            try:
                mixer.init()
                mixer.music.load("Incorrect.mp3")
                mixer.music.set_volume(1.0)
                mixer.music.play()
            except Exception as e:
                mixer.music.stop()

        def message_box(wording :list, value):
            try:
                messagebox.showinfo(title='Information', 
                    message=f'{wording[0]} {value}\n{wording[1]}')
            except:
                messagebox.showinfo(title='Information', 
                    message=f'เกิดข้อผิดพลาดกรุณาแจ้งหัวหน้างาน !!')

        def lockdown():
            self.txtSlot1.config(state="disabled")
            self.txtSlot2.config(state="disabled")
            self.txtStator.config(state="disabled")
            self.ClearBtn.config(state="!disabled")

        def unlock():
            self.txtSlot1.config(state="!disabled")
            self.txtSlot2.config(state="!disabled")
            self.txtStator.config(state="!disabled")
            self.ClearBtn.config(state="disabled")

        def checkEmpty():
            Alert : str = 'กรุณาตรวจสอบ\n'
            Check : bool = True

            if self.selected_Table.get() == "":
                Alert += 'หมายเลขโต๊ะทำงาน\n'
                Check = False

            if self.Arranger.get() == "":
                Alert += 'รหัสพนักงานของ(คนจัดงาน)\n'
                Check = False

            if self.Operator.get() == "":
                Alert += 'รหัสพนักงานของ(คนรันงาน)\n'
                Check = False

            if self.statorAssy.get() == "":
                Alert += 'Stator Assy Part No.\n'
                Check = False

            if self.slot1.get() == "":
                Alert += 'Slot1 Part No.\n'
                Check = False

            if self.slot2.get() == "":
                Alert += 'Slot2 Part No.\n'
                Check = False

            if len(self.tree.get_children()) <= 0:
                Alert += 'จำนวน Tray ในการรันงาน\n'
                Check = False

            if Check == False:
                messagebox.showinfo(
                    title='กรุณาระบุข้อมูลให้ครบถ้วน',
                    message = Alert,
                    icon = 'error'
                    )
            
            return Check
            
        
        ##################### ==== Add Data ==== ######################
        # For Add New Model Of Stator Assy from HB Division           #
        # into Database of this MechaII Program to correct components #
        ############################################################### 

        #### ==== Stator Assy ==== ####
        self.LblStator_Assy = ttk.Label(self.f1, text = 'Stator Assy : ', font=("Comic Sans MS", 14))
        self.LblStator_Assy.grid(row=0, column=0, padx=3, pady=3, sticky=tk.NE)

        self.Stator_Assy = tk.StringVar()
        self.txtStator_Assy = ttk.Entry(self.f1, textvariable=self.Stator_Assy, font=("Comic Sans MS", 14))
        self.txtStator_Assy.bind('<Return>', lambda event : onKeyPress_Stator_Assy(event, self.Stator_Assy.get()))
        self.txtStator_Assy.bind('<Tab>', lambda event : onKeyPress_Stator_Assy(event, self.Stator_Assy.get()))
        self.txtStator_Assy.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtStator_Assy,self.canvasStator_Assy))
        self.txtStator_Assy.grid(row=0, column=1, padx=3, pady=3, sticky=tk.EW)

        self.canvasStator_Assy = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasStator_Assy.grid(row=0, column=2, sticky=tk.EW)

        #### ==== Stator Assy SAP ==== ####
        self.LblStator_Assy_SAP = ttk.Label(self.f1, text = 'Stator Assy SAP : ', font=("Comic Sans MS", 14))
        self.LblStator_Assy_SAP.grid(row=0, column=3, padx=3, pady=3, sticky=tk.NE)

        self.Stator_Assy_SAP = tk.StringVar()
        self.txtStator_Assy_SAP = ttk.Entry(self.f1, textvariable=self.Stator_Assy_SAP, font=("Comic Sans MS", 14))
        self.txtStator_Assy_SAP.bind('<Return>', lambda event : onKeyPress_SAP(event, self.Stator_Assy_SAP.get(), self.txtStator_Assy_SAP, self.canvasStator_Assy_SAP, self.txtStatorStack))
        self.txtStator_Assy_SAP.bind('<Tab>', lambda event : onKeyPress_SAP(event, self.Stator_Assy_SAP.get(), self.txtStator_Assy_SAP, self.canvasStator_Assy_SAP, self.txtStatorStack))
        self.txtStator_Assy_SAP.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtStator_Assy_SAP, self.canvasStator_Assy_SAP))
        self.txtStator_Assy_SAP.grid(row=0, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasStator_Assy_SAP = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasStator_Assy_SAP.grid(row=0, column=5, sticky=tk.EW)

        #### ==== StatorStack ==== ####
        self.LblStatorStack = ttk.Label(self.f1, text = 'Stator Stack : ', font=("Comic Sans MS", 14))
        self.LblStatorStack.grid(row=1, column=0, padx=3, pady=3, sticky=tk.NE)

        self.StatorStack = tk.StringVar()
        self.txtStatorStack = ttk.Entry(self.f1, textvariable=self.StatorStack, font=("Comic Sans MS", 14))
        self.txtStatorStack.bind('<Return>', lambda event : onKeyPress_Stator_Stack(event, self.StatorStack.get()))
        self.txtStatorStack.bind('<Tab>', lambda event : onKeyPress_Stator_Stack(event, self.StatorStack.get()))
        self.txtStatorStack.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtStatorStack,self.canvasStatorStack))
        self.txtStatorStack.grid(row=1, column=1, padx=3, pady=3, sticky=tk.EW)

        self.canvasStatorStack = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasStatorStack.grid(row=1, column=2, sticky=tk.EW)

        #### ==== StatorStack SAP ==== ####
        self.LblStatorStack_SAP = ttk.Label(self.f1, text = 'Stator Stack SAP : ', font=("Comic Sans MS", 14))
        self.LblStatorStack_SAP.grid(row=1, column=3, padx=3, pady=3, sticky=tk.NE)

        self.StatorStack_SAP = tk.StringVar()
        self.txtStatorStack_SAP = ttk.Entry(self.f1, textvariable=self.StatorStack_SAP, font=("Comic Sans MS", 14))
        self.txtStatorStack_SAP.bind('<Return>', lambda event : onKeyPress_SAP(event, self.StatorStack_SAP.get(), self.txtStatorStack_SAP, self.canvasStatorStack_SAP, self.txtSlot_1))
        self.txtStatorStack_SAP.bind('<Tab>', lambda event : onKeyPress_SAP(event, self.StatorStack_SAP.get(), self.txtStatorStack_SAP, self.canvasStatorStack_SAP, self.txtSlot_1))
        self.txtStatorStack_SAP.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtStatorStack_SAP, self.canvasStatorStack_SAP))
        self.txtStatorStack_SAP.grid(row=1, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasStatorStack_SAP = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasStatorStack_SAP.grid(row=1, column=5, sticky=tk.EW)

        #### ==== Slot_1 ==== ####
        self.LblSlot_1 = ttk.Label(self.f1, text = 'Slot 1 : ', font=("Comic Sans MS", 14))
        self.LblSlot_1.grid(row=2, column=0, padx=3, pady=3, sticky=tk.NE)

        self.Slot_1 = tk.StringVar()
        self.txtSlot_1 = ttk.Entry(self.f1, textvariable=self.Slot_1, font=("Comic Sans MS", 14))
        self.txtSlot_1.bind('<Return>', lambda event : onKeyPress_Slot(event, self.Slot_1.get(), self.txtSlot_1, self.canvasSlot_1, self.txtSlot_1_SAP))
        self.txtSlot_1.bind('<Tab>', lambda event : onKeyPress_Slot(event, self.Slot_1.get(), self.txtSlot_1, self.canvasSlot_1, self.txtSlot_1_SAP))
        self.txtSlot_1.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtSlot_1,self.canvasSlot_1))
        self.txtSlot_1.grid(row=2, column=1, padx=3, pady=3, sticky=tk.EW)

        self.canvasSlot_1 = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasSlot_1.grid(row=2, column=2, sticky=tk.EW)

        #### ==== Slot_1 SAP ==== ####
        self.LblSlot_1_SAP = ttk.Label(self.f1, text = 'Slot 1 SAP : ', font=("Comic Sans MS", 14))
        self.LblSlot_1_SAP.grid(row=2, column=3, padx=3, pady=3, sticky=tk.NE)

        self.Slot_1_SAP = tk.StringVar()
        self.txtSlot_1_SAP = ttk.Entry(self.f1, textvariable=self.Slot_1_SAP, font=("Comic Sans MS", 14))
        self.txtSlot_1_SAP.bind('<Return>', lambda event : onKeyPress_SAP(event, self.Slot_1_SAP.get(), self.txtSlot_1_SAP, self.canvasSlot_1_SAP, self.txtSlot_2))
        self.txtSlot_1_SAP.bind('<Tab>', lambda event : onKeyPress_SAP(event, self.Slot_1_SAP.get(), self.txtSlot_1_SAP, self.canvasSlot_1_SAP, self.txtSlot_2))
        self.txtSlot_1_SAP.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtSlot_1_SAP, self.canvasSlot_1_SAP))
        self.txtSlot_1_SAP.grid(row=2, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasSlot_1_SAP = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasSlot_1_SAP.grid(row=2, column=5, sticky=tk.EW)

        #### ==== Slot_2 ==== ####
        self.LblSlot_2 = ttk.Label(self.f1, text = 'Slot 2 : ', font=("Comic Sans MS", 14))
        self.LblSlot_2.grid(row=3, column=0, padx=3, pady=3, sticky=tk.NE)

        self.Slot_2 = tk.StringVar()
        self.txtSlot_2 = ttk.Entry(self.f1, textvariable=self.Slot_2, font=("Comic Sans MS", 14))
        self.txtSlot_2.bind('<Return>', lambda event : onKeyPress_Slot(event, self.Slot_2.get(), self.txtSlot_2, self.canvasSlot_2, self.txtSlot_2_SAP))
        self.txtSlot_2.bind('<Tab>', lambda event : onKeyPress_Slot(event, self.Slot_2.get(), self.txtSlot_2, self.canvasSlot_2, self.txtSlot_2_SAP))
        self.txtSlot_2.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtSlot_2,self.canvasSlot_2))
        self.txtSlot_2.grid(row=3, column=1, padx=3, pady=3, sticky=tk.EW)

        self.canvasSlot_2 = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasSlot_2.grid(row=3, column=2, sticky=tk.EW)

        #### ==== Slot_2 SAP ==== ####
        self.LblSlot_2_SAP = ttk.Label(self.f1, text = 'Slot 2 SAP : ', font=("Comic Sans MS", 14))
        self.LblSlot_2_SAP.grid(row=3, column=3, padx=3, pady=3, sticky=tk.NE)

        self.Slot_2_SAP = tk.StringVar()
        self.txtSlot_2_SAP = ttk.Entry(self.f1, textvariable=self.Slot_2_SAP, font=("Comic Sans MS", 14))
        self.txtSlot_2_SAP.bind('<Return>', lambda event : onKeyPress_SAP(event, self.Slot_2_SAP.get(), self.txtSlot_2_SAP, self.canvasSlot_2_SAP))
        self.txtSlot_2_SAP.bind('<Tab>', lambda event : onKeyPress_SAP(event, self.Slot_2_SAP.get(), self.txtSlot_2_SAP, self.canvasSlot_2_SAP))
        self.txtSlot_2_SAP.bind("<Button-1>", lambda event: focus_in_add_panel(event, self.txtSlot_2_SAP, self.canvasSlot_2_SAP))
        self.txtSlot_2_SAP.grid(row=3, column=4, padx=3, pady=3, sticky=tk.EW)

        self.canvasSlot_2_SAP = tk.Canvas(self.f1, width=self.w, height=self.w)
        self.canvasSlot_2_SAP.grid(row=3, column=5, sticky=tk.EW)

        self.Btn_Add = ttk.Button(self.f1, text='เพิ่มข้อมูล (Add Data)', command = lambda:onClick_AddData())
        self.Btn_Add.config(state="disabled")
        self.Btn_Add.grid(row=4, column=4, padx=3, pady=3, sticky=tk.NE)

        def focus_in_add_panel(event, Txt : ttk.Entry, Canvas : tk.Canvas = None):
            try:
                Txt.delete(0,END)
                Canvas.delete('all')
                Txt.focus()
            except:
                pass

        def onKeyPress_Stator_Assy(event, value):
            # https://www.programiz.com/python-programming/regex
            x = re.match("^16..-", value)

            if len(value) > 6 and x:
                createGreenLight(self.canvasStator_Assy)
                self.txtStator_Assy_SAP.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                Thread(focus_in_add_panel(event,self.txtStator_Assy,self.canvasStator_Assy)).start()
                return

        def onKeyPress_Stator_Stack(event, value):
            x = re.match("^19..-", value)

            if len(value) > 6 and x:
                createGreenLight(self.canvasStatorStack)
                self.txtStatorStack_SAP.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                Thread(focus_in_add_panel(event,self.txtStatorStack,self.canvasStatorStack)).start()
                return

        def onKeyPress_Slot(event, value, tx: ttk.Entry ,cnv : tk.Canvas, tx_next : ttk.Entry):
            x = re.match("^72..-", value)

            if len(value) > 6 and x:
                createGreenLight(cnv)
                tx_next.focus()
            else:
                Thread(correct()).start()
                Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                Thread(focus_in_add_panel(event, tx, cnv)).start()
                return

        def onKeyPress_SAP(event, value , tx: ttk.Entry ,cnv : tk.Canvas, tx_next : ttk.Entry = None):
            x = re.match("^1[0-9]{10}", value)

            if x or value == '-':
                createGreenLight(cnv)
                if not tx_next == None:
                    tx_next.focus()
                else:
                    checkCondition()
            else:
                Thread(correct()).start()
                Thread(message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],value)).start()
                Thread(focus_in_add_panel(event, tx, cnv)).start()
                return

        def checkCondition():
            check = False
            check_SAP = False

            a_0 = self.Stator_Assy.get()
            a = a_0[2:4]
            b_0 = self.StatorStack.get()
            b = b_0[2:4]
            c_0 = self.Slot_1.get()
            c = c_0[2:4]
            d_0 = self.Slot_2.get()
            d = d_0[2:4]

            e = set()

            e_0 = self.Stator_Assy_SAP.get()
            if e_0 != None and e_0 != '-':
                e.add(e_0)

            f_0 = self.StatorStack_SAP.get()
            if f_0 != None and f_0 != '-':
                e.add(f_0)

            g_0 = self.Slot_1_SAP.get()
            if g_0 != None and g_0 != '-':
                e.add(g_0)

            h_0 = self.Slot_2_SAP.get()
            if h_0 != None and h_0 != '-':
                e.add(h_0)

            if a==b and a==c and a==d and b==c and b==d and c==d :
                check = True

            check_SAP = True if len(e) == len(set(e)) else False #Check Duplicates of List using set

            if check and check_SAP :
                self.Btn_Add.config(state="!disabled")
            else:
                message_box(['Format Not Correct !!! (ข้อมูลผิดพลาด)','กรุณาตรวจสอบความถูกต้อง'],"ข้อมูล part no. / sap no.")


        def ClearAll_AddData():
            self.txtStator_Assy.delete(0,END)
            self.txtStator_Assy_SAP.delete(0,END)
            self.txtStatorStack.delete(0,END)
            self.txtStatorStack_SAP.delete(0,END)
            self.txtSlot_1.delete(0,END)
            self.txtSlot_1_SAP.delete(0,END)
            self.txtSlot_2.delete(0,END)
            self.txtSlot_2_SAP.delete(0,END)
            self.canvasStator_Assy.delete('all')
            self.canvasStator_Assy_SAP.delete('all')
            self.canvasStatorStack.delete('all')
            self.canvasStatorStack_SAP.delete('all')
            self.canvasSlot_1.delete('all')
            self.canvasSlot_1_SAP.delete('all')
            self.canvasSlot_2.delete('all')
            self.canvasSlot_2_SAP.delete('all')
            self.Btn_Add.config(state="disabled")

        def onClick_AddData():

            MsgBox = messagebox.askquestion ('Add Data',f"""
            Are you sure to Add\n
            Stator Assy No.\t:\t{self.Stator_Assy.get()}\n
            Stator Assy SAP\t:\t{self.Stator_Assy_SAP.get()}\n
            Stator Stack No.\t:\t{self.StatorStack.get()}\n
            Stator Stack SAP\t:\t{self.StatorStack_SAP.get()}\n
            Slot 1\t:\t{self.Slot_1.get()}\n
            Slot 1 SAP\t:\t{self.Slot_1_SAP.get()}\n
            Slot 2\t:\t{self.Slot_2.get()}\n
            Slot 2 SAP\t:\t{self.Slot_2_SAP.get()}\n
            """,icon='warning')

            if MsgBox == 'yes':
                addmdl = StatorAssyDetail()
                addmdl.setStatorAssy(self.Stator_Assy.get())
                addmdl.setNewSAP(self.Stator_Assy_SAP.get())
                addmdl.setStackNo(self.StatorStack.get())
                addmdl.setStackSAP(self.StatorStack_SAP.get())
                addmdl.setSlot_1(self.Slot_1.get())
                addmdl.setSlot_1_SAP(self.Slot_1_SAP.get())
                addmdl.setSlot_2(self.Slot_2.get())
                addmdl.setSlot_2_SAP(self.Slot_2_SAP.get())

                self.ctrl.insertVaribleIntoTable(
                    addmdl.getNewSAP(),
                    addmdl.getStatorAssy(),
                    addmdl.getStackNo(),
                    addmdl.getStackSAP(),
                    addmdl.getSlot_1(),
                    addmdl.getSlot_1_SAP(),
                    addmdl.getSlot_2(),
                    addmdl.getSlot_2_SAP()
                )
                messagebox.showinfo(title='Information', message=f'New Data {self.Stator_Assy.get()} was Added')
            else:
                pass

            ClearAll_AddData()

        ##################### ==== Edit Data ==== #####################
        # For Edit Details and update following HB                    #
        # into Database of this MechaII Program to correct components #
        ###############################################################
        
        self.LbFEditStator_AaddSAP = LabelFrame(self.f2, text='Stator Assy update SAP No.')
        self.LbFEditStator_AaddSAP.grid(row=0, column=0, padx=2, pady=2, sticky=tk.NSEW)

        self.LblStator_ASearch = Label(self.LbFEditStator_AaddSAP, text='Stator Assy No.', font=("Comic Sans MS", 14))
        self.LblStator_ASearch.grid(row=0, column=0, padx=2, pady=2, sticky=tk.NW)

        self.Stator_ASearch = tk.StringVar()
        self.TxtStator_ASearch = ttk.Entry(self.LbFEditStator_AaddSAP, textvariable=self.Stator_ASearch, font=("Comic Sans MS", 14))
        self.TxtStator_ASearch.bind('<Return>',
            lambda event : Update_Click(event, ['Statorassy'], self.Stator_ASearch.get(), [self.LblStator_A_SAP_Find,self.txtStator_A_SAP_Find,self.Btn_StatorA_Confirm]))
        self.TxtStator_ASearch.bind('<Tab>',
            lambda event : Update_Click(event, ['Statorassy'], self.Stator_ASearch.get(), [self.LblStator_A_SAP_Find,self.txtStator_A_SAP_Find,self.Btn_StatorA_Confirm]))
        self.TxtStator_ASearch.bind("<Button-1>",
            lambda event: focus_in_add_panel_edit(event, self.TxtStator_ASearch, [self.LblStator_A_SAP_Find,self.txtStator_A_SAP_Find,self.Btn_StatorA_Confirm]))
        self.TxtStator_ASearch.grid(row=0, column=1, padx=2, pady=2, sticky=tk.NW)

        ############## --------------- SQL --------------- ##############

        self.LblStator_A_SAP_Find = Label(self.LbFEditStator_AaddSAP, text='SAP No. of Stator Assy', font=("Comic Sans MS", 14))
        self.LblStator_A_SAP_Find.grid(row=0, column=3, padx=2, pady=2, sticky=tk.NW)
        self.LblStator_A_SAP_Find['state'] = 'disabled'

        self.Stator_A_SAP_Find = tk.StringVar()
        self.txtStator_A_SAP_Find = ttk.Entry(self.LbFEditStator_AaddSAP, textvariable=self.Stator_A_SAP_Find, font=("Comic Sans MS", 14))
        self.txtStator_A_SAP_Find.bind('<Return>', lambda event : print('test'))
        self.txtStator_A_SAP_Find.bind('<Tab>', lambda event : print('test'))
        self.txtStator_A_SAP_Find.bind("<Button-1>", lambda event: print('test'))
        self.txtStator_A_SAP_Find.grid(row=0, column=4, padx=2, pady=2, sticky=tk.NW)
        self.txtStator_A_SAP_Find['state'] = 'disabled'

        self.Btn_StatorA_Confirm = Button(self.LbFEditStator_AaddSAP, text='Confirm SAP No.')
        self.Btn_StatorA_Confirm.grid(row=0, column=5, padx=2, pady=2, sticky=tk.NSEW)
        self.Btn_StatorA_Confirm['state'] = 'disabled'

        ####################################################################################################

        self.LbFEditInsulatorSAP = LabelFrame(self.f2, text='Insulator update SAP No.')
        self.LbFEditInsulatorSAP.grid(row=1, column=0, padx=2, pady=2, sticky=tk.NSEW)

        self.LblInsulatorSearch = Label(self.LbFEditInsulatorSAP, text='Insulator No.', font=("Comic Sans MS", 14))
        self.LblInsulatorSearch.grid(row=0, column=0, padx=2, pady=2, sticky=tk.NW)

        self.InsulatorSearch = tk.StringVar()
        self.TxtInsulatorSearch = ttk.Entry(self.LbFEditInsulatorSAP, textvariable=self.InsulatorSearch, font=("Comic Sans MS", 14))
        self.TxtInsulatorSearch.bind('<Return>',
            lambda event : Update_Click(event, ['Slot_1','Slot_2'], self.InsulatorSearch.get(), [self.LblInsulator_SAP_Find, self.TxtInsulator_SAP_Find, self.Btn_Insulator_Confirm]))
        self.TxtInsulatorSearch.bind('<Tab>',
            lambda event : Update_Click(event, ['Slot_1','Slot_2'], self.InsulatorSearch.get(), [self.LblInsulator_SAP_Find, self.TxtInsulator_SAP_Find, self.Btn_Insulator_Confirm]))
        self.TxtInsulatorSearch.bind("<Button-1>",
            lambda event : focus_in_add_panel_edit(event, self.TxtInsulatorSearch, [self.LblInsulator_SAP_Find, self.TxtInsulator_SAP_Find, self.Btn_Insulator_Confirm]))
        self.TxtInsulatorSearch.grid(row=0, column=1, padx=2, pady=2, sticky=tk.NW)

        ############## --------------- SQL --------------- ##############

        self.LblInsulator_SAP_Find = Label(self.LbFEditInsulatorSAP, text='SAP No. of Insulator', font=("Comic Sans MS", 14))
        self.LblInsulator_SAP_Find.grid(row=0, column=3, padx=2, pady=2, sticky=tk.NW)
        self.LblInsulator_SAP_Find['state'] = 'disabled'

        self.Insulator_SAP_Find = tk.StringVar()
        self.TxtInsulator_SAP_Find = ttk.Entry(self.LbFEditInsulatorSAP, textvariable=self.Stator_A_SAP_Find, font=("Comic Sans MS", 14))
        self.TxtInsulator_SAP_Find.bind('<Return>', lambda event : print('test'))
        self.TxtInsulator_SAP_Find.bind('<Tab>', lambda event : print('test'))
        self.TxtInsulator_SAP_Find.bind("<Button-1>", lambda event: print('test'))
        self.TxtInsulator_SAP_Find.grid(row=0, column=4, padx=2, pady=2, sticky=tk.NW)
        self.TxtInsulator_SAP_Find['state'] = 'disabled'

        self.Btn_Insulator_Confirm = Button(self.LbFEditInsulatorSAP, text='Confirm SAP No.')
        self.Btn_Insulator_Confirm.grid(row=0, column=5, padx=2, pady=2, sticky=tk.NSEW)
        self.Btn_Insulator_Confirm['state'] = 'disabled'


        ####################################################################################################

        self.LbFEditStatorChg_SAP = LabelFrame(self.f2, text='Stator update SAP No.')
        self.LbFEditStatorChg_SAP.grid(row=2, column=0, padx=2, pady=2, sticky=tk.NSEW)

        self.LblStatorSearch = Label(self.LbFEditStatorChg_SAP, text='Stator No.', font=("Comic Sans MS", 14))
        self.LblStatorSearch.grid(row=0, column=0, padx=2, pady=2, sticky=tk.NW)

        self.StatorSearch = tk.StringVar()
        self.TxtStatorSearch = ttk.Entry(self.LbFEditStatorChg_SAP, textvariable=self.StatorSearch, font=("Comic Sans MS", 14))
        self.TxtStatorSearch.bind('<Return>', 
            lambda event : Update_Click(event, ['StackNo'], self.StatorSearch.get(), [self.LblStator_SAP_Find, self.TxtStator_SAP_Find, self.Btn_Stator_Confirm]))
        self.TxtStatorSearch.bind('<Tab>',
            lambda event : Update_Click(event, ['StackNo'], self.StatorSearch.get(), [self.LblStator_SAP_Find, self.TxtStator_SAP_Find, self.Btn_Stator_Confirm]))
        self.TxtStatorSearch.bind("<Button-1>",
            lambda event : focus_in_add_panel_edit(event, self.TxtStatorSearch, [self.LblStator_SAP_Find, self.TxtStator_SAP_Find, self.Btn_Stator_Confirm]))
        self.TxtStatorSearch.grid(row=0, column=1, padx=2, pady=2, sticky=tk.NW)

        ############## --------------- SQL --------------- ##############

        self.LblStator_SAP_Find = Label(self.LbFEditStatorChg_SAP, text='SAP No. of Stator', font=("Comic Sans MS", 14))
        self.LblStator_SAP_Find.grid(row=0, column=3, padx=2, pady=2, sticky=tk.NW)
        self.LblStator_SAP_Find['state'] = 'disabled'

        self.Stator_SAP_Find = tk.StringVar()
        self.TxtStator_SAP_Find = ttk.Entry(self.LbFEditStatorChg_SAP, textvariable=self.Stator_A_SAP_Find, font=("Comic Sans MS", 14))
        self.TxtStator_SAP_Find.bind('<Return>', lambda event : print('test'))
        self.TxtStator_SAP_Find.bind('<Tab>', lambda event : print('test'))
        self.TxtStator_SAP_Find.bind("<Button-1>", lambda event: print('test'))
        self.TxtStator_SAP_Find.grid(row=0, column=4, padx=2, pady=2, sticky=tk.NW)
        self.TxtStator_SAP_Find['state'] = 'disabled'

        self.Btn_Stator_Confirm = Button(self.LbFEditStatorChg_SAP, text='Confirm SAP No.')
        self.Btn_Stator_Confirm.grid(row=0, column=5, padx=2, pady=2, sticky=tk.NSEW)
        self.Btn_Stator_Confirm['state'] = 'disabled'

        def Update_Click(event, find_col, find, element:list):
            x = self.ctrl.select_count(find_col, find)
            if x > 0:
                for i in element:
                    i.configure(state='active')
            else: 
                pass

        def focus_in_add_panel_edit(event, Txt:ttk.Entry, element:list):
            try:
                Txt.delete(0,END)
                Txt.focus()

                for i in element:
                    i.configure(state='disabled')
            except:
                pass