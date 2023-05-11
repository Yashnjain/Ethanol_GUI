from email.mime import message
import tkinter as tk
from tkinter.filedialog import askdirectory, askopenfilename
from tkinter import N, Menubutton, Tk, StringVar, Text
from tkinter import PhotoImage
from tkinter.font import Font
from tkinter.ttk import Label
from tkinter import Button
from tkinter.ttk import Frame, Style
from tkinter.ttk import OptionMenu
from tkinter import Label as label
from tkcalendar import DateEntry
from tkinter import messagebox
# from typing import Text
import traceback
from pandas.core import frame 
import requests, json
from datetime import date, datetime, timedelta
import numpy as np
import glob, time
from tkinter.messagebox import showerror
import pandas as pd
import os
import xlwings as xw
from tabula import read_pdf
# import PyPDF2
from collections import defaultdict
import xlwings.constants as win32c
import sys, traceback
import PyPDF2
from collections import OrderedDict
import calendar
from dateutil.relativedelta import relativedelta
import shutil
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.options import Options
import re
import array
from CASH.cash import cash
from NLV_FUTURES.NLV_FUTURES import NLV_FUTURESSS
from FOB.fob import fob_runner
from AR_ageing_bulk.ar_ageing_bulk import ar_ageing_bulk
from ar_ageing_rack.ar_ageing_rack import ar_ageing_rack
from purchased_ar.purchased_ar import purchased_ar
from unbilled_ar.unbilled_ar import unbilled_ar
from RackBackTrack.rackbacktrack import rackbacktrack

from Open_GR.open_gr import openGr
# from Common.common import set_borders,freezepanes_for_tab,interior_coloring,conditional_formatting2,interior_coloring_by_theme,num_to_col_letters,insert_all_borders,conditional_formatting,knockOffAmtDiff,row_range_calc,thick_bottom_border



# path = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\Revelio'
path = r'J:\India\BBR\IT_BBR\Reports\Ethanol_gui'
today = datetime.strftime(date.today(), format = "%d%m%Y")


root = Tk()
root.title('ETHANOL APP')
root.geometry('648x696')
photo = PhotoImage(file = path + '\\'+'biourjaLogo.png')
root.iconphoto(False, photo)
root["bg"]= "white"


frame_title = Frame(root)
frame_options = Frame(root)
frame_folder = Frame(root)
frame_submit = Frame(root)
frame_msg = Frame(root)
s = Style(frame_options)
s.configure("TMenubutton", background="#f5fcfc",width=19, font=("Book Antiqua", 12))
s.configure("TMenu", width=19)
s.configure("TFrame", background="white")


class MyDateEntry(DateEntry):
    def __init__(self, master=None, **kw):
        DateEntry.__init__(self, master=master, date_pattern='mm.dd.yyyy',**kw)
        # add black border around drop-down calendar
        self._top_cal.configure(bg='black', bd=1)
        # add label displaying today's date below
        label(self._top_cal, bg='gray90', anchor='w',
                 text='Today: %s' % date.today().strftime('%x')).pack(fill='both', expand=1)


def open_gr(input_date,output_date):
    try:
        msg = openGr(input_date, output_date)
        return msg
    except Exception as e:
        raise e

def unbilled_ar_(input_date, output_date):
    try:
        msg = unbilled_ar(input_date, output_date)
        return msg
    except Exception as e:
        raise e

def purchased_ar_(input_date, output_date):
    try:
        msg = purchased_ar(input_date, output_date)
        return msg
    except Exception as e:
        raise e


def ar_ageing_bulk_(input_date, output_date):
    try:
        msg = ar_ageing_bulk(input_date, output_date)
        return msg
    except Exception as e:
        raise e

def ar_ageing_rack_(input_date, output_date):
    try:
        msg = ar_ageing_rack(input_date, output_date)
        return msg
    except Exception as e:
        raise e


def bbr_nlv_futures(start_date,end_date):
    try:
        msg = NLV_FUTURESSS(start_date,end_date)
        return msg
    except Exception as e:
        raise e

def bbr_cash(start_date,end_date):
    try:
        msg = cash(start_date,end_date)
        return msg
    except Exception as e:
        raise e
def bbr_fob(start_date,end_date):
    try:
        msg = fob_runner(start_date,end_date)
        return msg
    except Exception as e:
        raise e
    
def call_rackbacktrack(start_date,end_date):
    try:
        msg = rackbacktrack(start_date,end_date)
        return msg
    except Exception as e:
        raise e
        
def main():
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            
            root.destroy()
            sys.exit()
    def callback_2():
    
        # def report_callback_exception(self, exc, val, tb):
        #     showerror("Error", message=str(exc) + str(val) +str(tb))

        # try:
        if submit_text.get() != "Started" and 'Select' not in Rep_variable.get():
            submit_text.set("Started")
            text_box.delete(1.0, "end")
            text_box.tag_configure("center", justify='center')
            text_box.tag_add("center", 6.0, "end")
            text_box.insert("end", f"In Process", "center")
            root.update()
            
            print(inp_date.get())
            print(Rep_variable.get())
            input_date = inp_date.get()
            output_date = out_date.get()
            func_to_call = Rep_variable.get()
            msg = wp_job_ids[func_to_call](input_date, output_date)
            text_box.delete(1.0, "end")
            text_box.insert("end", f"\n{msg}", "center")
            submit_text.set("Submit")
            Rep_variable.set('Select')
            root.update()

            print()
        
        elif 'Select' in Rep_variable.get():
            text_box.delete(1.0, "end")
            text_box.insert("end", f"Please select job first", "center")


        root.update()
        # except Exception as e:
        #     Tk.report_callback_exception = report_callback_exception
        
        
    # def callback():
    #     try:
    #         threading.Thread(target=callback_2).start()
    #     except Exception as e:
    #         raise e
        
        
    def report_callback_exception(self, exc, val, tb):
        msg = traceback.format_exc()
        showerror("Error", message=msg)
        text_box.delete(1.0, "end")
        text_box.insert("end", str(exc), "center")
        submit_text.set("Submit")
        Rep_variable.set('Select')
        root.update()

    Tk.report_callback_exception = report_callback_exception
    frame_title.grid(row=0, column=1,pady=(24,0),columnspan=3, padx=(30,0))
    logo = PhotoImage(file = path + '\\'+'Ethanol_Logo.png')
    # logo = PhotoImage(file = path + '\\'+'wp_logo.png')


    title = Label(frame_title,image=logo, background='white')
    # title = Label(frame_title, text="Revelio Report Generator", font=("Algerian", 28), foreground='black', background="white")
    title.grid(column=1,row=0)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    # input_date=None
    # output_date = None
    frame_options.grid(row=1,column=0, pady=30, padx=35, columnspan=2, rowspan=3)
    wp_job_ids = {'ABS':1,'Purchased AR Report':purchased_ar,'Ar Ageing Report(Bulk)':ar_ageing_bulk, 'Open Gr':open_gr ,
    'Ar Ageing Report(Rack)':ar_ageing_rack,'Unbilled AR Report':unbilled_ar,'Cash BBR':bbr_cash,'NLV BBR':bbr_nlv_futures, 'Rack Back Track':call_rackbacktrack,'BBR FOB':bbr_fob}
    # wp_job_ids = {'ABS':1,'BBR':bbr,'CPR Report':cpr, 'Freight analysis':freight_analysis, 'CTM combined':ctm,'MTM Report':mtm_report,
    #                 'MOC Interest Allocation':moc_interest_alloc,'Open AR':open_ar,'Open AP':open_ap, 'Unsettled Payable Report':unsetteled_payables,'Unsettled Receivable Report':unsetteled_receivables,
    #                 'Storage Month End Report':strg_month_end_report, "Month End BBR":bbr_monthEnd, "Bank Recons Report":bank_recons_rep}
    # department_ids = {'Select \t\t\t\t\t\t\t\t\t': 9, 'Ethanol\t\t\t\t\t\t\t\t': 1, 'WestPlains': 8}
    Rep_variable = StringVar()
    doc_type_variable = StringVar()
    doc_type_variable.set('Select')
    folderPath = StringVar()
    macroPath = StringVar()
    # Dep_variable.trace('w', update_options_B)
    dep_label = Label(frame_options, text='Select Job:                  ', font=("Book Antiqua bold", 16), foreground="#FF0000", background="white")
    dep_label.grid(row=0, column=0)
    Dep_option = OptionMenu(frame_options, Rep_variable, *wp_job_ids.keys())
    
    Dep_option["menu"].configure(background="white", font=("Arial", 12)) #, bg='#20bebe', fg='white')
    # Dep_option["menu"].config(width=19)
    Dep_option.grid(row=0, column=1)
    Rep_variable.set('Select \t\t\t\t\t\t\t\t\t')
    blank = Label(frame_options, text="                                ", font=("Helvetica", 16), foreground="blue", justify='left', background="white")
    blank.grid(row=1, column=0)
    blank2 = Label(frame_options, text="             ", font=("Helvetica", 16), foreground="green", justify='left', background="white")
    blank2.grid(row=1, column=1)
    # doc_label = Label(frame_options, text="Select Doc_Type:     ", font=("Book Antiqua bold", 16), foreground="#ff8c00", background="white")
    # doc_label.grid(row=2, column=0)
    # doc_type_option = OptionMenu(frame_options, doc_type_variable, '')
    # doc_type_option["menu"].configure(background="white", font=("Arial", 12))
    # doc_type_option.grid(row=2, column=1)

    blank3 = Label(frame_options, text="                                ", font=("Helvetica", 16), foreground="blue", justify='left', background="white")
    blank3.grid(row=3, column=0)
    folder_label = Label(frame_options, text="Select Input Date:     ", font=("Book Antiqua bold", 16), foreground="#FF0000", background="white",justify='left')
    folder_label.grid(row=4, column=0)
    browse_text = StringVar()
    inp_date = MyDateEntry(master=frame_options, width=17, selectmode='day') #Button(frame_options, textvariable=browse_text, command=getFolderPath, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, activebackground="#20bebb")
    browse_text.set("Browse")
    inp_date.grid(row=4, column=1)

    blank4 = Label(frame_options, text="                                ", font=("Helvetica", 16), foreground="blue", justify='left', background="white")
    blank4.grid(row=5, column=0)
    macro_label = Label(frame_options, text="Select Prev File Date:", font=("Book Antiqua bold", 16), foreground="#FF0000", background="white",justify='left')
    macro_label.grid(row=6, column=0)
    browse_text2 = StringVar()
    out_date = MyDateEntry(master=frame_options, width=17, selectmode='day') #Button(frame_options, textvariable=browse_text2, command=getFilePath, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, activebackground="#20bebb")
    browse_text2.set("Browse")
    out_date.grid(row=6, column=1)

    frame_folder.grid(row=2, column=2, padx=(28,0))
    

    frame_submit.grid(row=5, column=1,columnspan=3)
    submit_text = StringVar()
    submit = Button(frame_submit, textvariable=submit_text, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, command=callback_2, activebackground="#20bebb")
    submit.grid(row=0, column=1, padx=(30,0))
    submit_text.set("Submit")
    
    # if doc_type_variable.get() == "Select \t\t\t\t\t\t\t\t\t":
    #     sel_Folder["state"] = "disabled"
    #     submit["state"] = "disabled"
        

    # text_box = Text(root, height=10, width=50, padx=15, pady=15)
    # text_box.insert(1.0, "Select Details, and click Select folder n Submit")
    # text_box.tag_configure("center", justify="center")
    # text_box.tag_add("center", 1.0), "end"
    # text_box.grid(column=1, row=6)
    blank3 = Label(frame_submit, text="             ", font=("Helvetica", 16), foreground="green", justify='left', background="white")
    blank3.grid(row=1, column=1)
    
    
    staus_text = StringVar()
    frame_msg.grid(row=7,column=1,columnspan=3) ##, padx=(180,0))
    text_box = Text(frame_msg, background="white",font=("Raleway", 10), width=88, height=10, borderwidth=0)

    # label_2 = Label(root, textvariable=label_2_text, background="white", justify='center',font=("Raleway", 12)) 
    text_box.grid(row=7, column=1,columnspan=3, padx=(14,0)) # column
    # label_2.grid(row=8, column=1,columnspan=2)
    # 
    # label_2_text.set("")

    root.mainloop()


if __name__ == '__main__':
    main()

