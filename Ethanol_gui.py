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
from NLV_FUTURES.NLV_FUTURES import NLV_FUTURES

from Open_GR.open_gr import openGr
from Common.common import set_borders,freezepanes_for_tab,interior_coloring,conditional_formatting2,interior_coloring_by_theme,num_to_col_letters,insert_all_borders,conditional_formatting,knockOffAmtDiff,row_range_calc,thick_bottom_border



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

# def set_borders(border_range):
#     for border_id in range(7,13):
#         border_range.api.Borders(border_id).LineStyle=1
#         border_range.api.Borders(border_id).Weight=2


# def freezepanes_for_tab(cellrange:str,working_sheet,working_workbook):
#     try:
#         working_sheet.activate()
#         working_sheet.api.Rows(cellrange).Select()
#         working_workbook.app.api.ActiveWindow.FreezePanes = True
#     except Exception as e:
#         raise e        

# def interior_coloring(colour_value,cellrange:str,working_sheet,working_workbook):
#     try:
#         working_sheet.activate()
#         if working_sheet.api.AutoFilterMode:
#             working_sheet.api.Range(cellrange).SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         else:
#             working_sheet.api.Range(cellrange).Select()
#         a = working_workbook.app.selection.api.Interior
#         a.Pattern = win32c.Constants.xlSolid
#         a.PatternColorIndex = win32c.Constants.xlAutomatic
#         a.Color = colour_value
#         a.TintAndShade = 0
#         a.PatternTintAndShade = 0        
#     except Exception as e:
#         raise e  

# def conditional_formatting2(range:str,working_sheet,working_workbook):
#     try:
#         font_colour = -16383844
#         Interior_colour = 13551615
#         working_sheet.api.Range(range).Select()
#         working_workbook.app.selection.api.FormatConditions.Add(Type:=win32c.FormatConditionType.xlCellValue, Operator:=win32c.FormatConditionOperator.xlLess,Formula1:="=0")
#         working_workbook.app.selection.api.FormatConditions(working_workbook.app.selection.api.FormatConditions.Count).SetFirstPriority()
#         working_workbook.app.selection.api.FormatConditions(1).Font.Color = font_colour
#         working_workbook.app.selection.api.FormatConditions(1).Interior.Color = Interior_colour
#         working_workbook.app.selection.api.FormatConditions(1).Interior.PatternColorIndex = win32c.Constants.xlAutomatic
#         working_workbook.app.selection.api.FormatConditions(1).StopIfTrue = False
#         return font_colour,Interior_colour
#     except Exception as e:
#         raise e

# def interior_coloring_by_theme(pattern_tns,tintandshade,colour_value,cellrange:str,working_sheet,working_workbook):
#     try:
#         working_sheet.activate()
#         if working_sheet.api.AutoFilterMode:
#             working_sheet.api.Range(cellrange).SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         else:
#             working_sheet.api.Range(cellrange).Select()
#         a = working_workbook.app.selection.api.Interior
#         # a.Pattern = win32c.Constants.xlSolid
#         a.PatternColorIndex = win32c.Constants.xlAutomatic
#         a.ThemeColor = colour_value
#         a.TintAndShade = tintandshade
#         a.PatternTintAndShade = pattern_tns    
#     except Exception as e:
#         raise e  
            

# def num_to_col_letters(num):
#     try:
#         letters = ''
#         while num:
#             mod = (num - 1) % 26
#             letters += chr(mod + 65)
#             num = (num - 1) // 26
#         return ''.join(reversed(letters))
#     except Exception as e:
#         raise e

# def insert_all_borders(cellrange:str,working_sheet,working_workbook):
#         working_sheet.api.Range(cellrange).Select()
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
#         linestylevalues=[win32c.BordersIndex.xlEdgeLeft,win32c.BordersIndex.xlEdgeTop,win32c.BordersIndex.xlEdgeBottom,win32c.BordersIndex.xlEdgeRight,win32c.BordersIndex.xlInsideVertical,win32c.BordersIndex.xlInsideHorizontal]
#         for values in linestylevalues:
#             a=working_workbook.app.selection.api.Borders(values)
#             a.LineStyle = win32c.LineStyle.xlContinuous
#             a.ColorIndex = 0
#             a.TintAndShade = 0
#             a.Weight = win32c.BorderWeight.xlThin

# def conditional_formatting(range:str,working_sheet,working_workbook):
#     try:
#         working_workbook.activate()
#         working_sheet.activate()
#         font_colour = -16383844
#         Interior_colour = 13551615
#         working_sheet.api.Range(range).Select()
#         working_workbook.app.selection.api.FormatConditions.AddUniqueValues()
#         working_workbook.app.selection.api.FormatConditions(working_workbook.app.selection.api.FormatConditions.Count).SetFirstPriority()

#         working_workbook.app.selection.api.FormatConditions(1).DupeUnique = win32c.DupeUnique.xlDuplicate

#         working_workbook.app.selection.api.FormatConditions(1).Font.Color = font_colour
#         working_workbook.app.selection.api.FormatConditions(1).Interior.Color = Interior_colour
#         working_workbook.app.selection.api.FormatConditions(1).Interior.PatternColorIndex = win32c.Constants.xlAutomatic
#         return font_colour,Interior_colour
#     except Exception as e:
#         raise e

# def knockOffAmtDiff(curr,final, wb, input_sht, input_sht2, credit_col_letter, debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter, row_dict, eth_trueup_col_letter=None):
#     try:
#         print(row_dict["Knock_Off"])
#         if abs(input_sht.range(f"{credit_col_letter}{curr}").value) == abs(input_sht2.range(f"{debit_col_letter}{final}").value):
#             print(f"Moving {curr} to knockoff")
#             knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row

#             #Copy Pasting Whole data
#             # input_sht.range(f"{curr}:{final}").api.Copy()
#             # wb.activate()
#             # knock_off_sht.activate()
#             # knock_off_sht.range(f"A{knock_off_last_row+1}").api.Select()
#             # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#             # knock_off_sht.autofit()
#             if input_sht==input_sht2:
#                 # input_sht.range(f"{curr}:{final}").copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))

#                 # input_sht.range(f"{curr}:{final}").delete()
#                 # input_sht.range(f"{curr}:{final}").color ="#00FF00"
                
#                 if not len(row_dict["Knock_Off"]):
#                     row_dict["Knock_Off"] = [[f"{curr}:{final}"]]
#                     # knockoff_list.append(f"{curr}:{final}")
#                 # elif int(knockoff_list[-1].split(":")[-1]) == curr-1:   #prev final == currnt -1
#                 elif len(row_dict["Knock_Off"][-1]) <=24:
#                     if int(row_dict["Knock_Off"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1
#                         # knockoff_list[-1] = f'{knockoff_list[-1].split(":")[0]}:{final}'
#                         row_dict["Knock_Off"][-1][-1] = f'{row_dict["Knock_Off"][-1][-1].split(":")[0]}:{final}'
#                     else:
#                         # knockoff_list.append(f"{curr}:{final}")
#                         row_dict["Knock_Off"][-1].append(f"{curr}:{final}")
#                 elif len(row_dict["Knock_Off"][-1]) >24:
#                     row_dict["Knock_Off"].append([f"{curr}:{final}"])
                
#             else:
#                 input_sht.range(f"{curr}:{curr}").copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))
#                 input_sht2.range(f"B{final}:{eth_trueup_col_letter}{final}").copy(knock_off_sht.range(f"A{knock_off_last_row+2}"))

#                 #shifting credit amount to right cell copied from ethanol accrual
#                 knock_off_sht.range(f"K{knock_off_last_row+2}").copy(knock_off_sht.range(f"L{knock_off_last_row+2}"))
#                 knock_off_sht.range(f"K{knock_off_last_row+2}").clear()
#                 knock_off_sht.range(f"M{knock_off_last_row+2}").clear()#Clearing Final Amount

#                 input_sht.range(f"{curr}:{curr}").delete()
#                 input_sht2.range(f"{final}:{final}").delete()
#                 curr-=1
                

#                 # input_sht.range(f"{curr}:{curr}").delete()
#                 # input_sht.range(f"{curr}:{curr}").color ="#00FF00"
#                 # input_sht2.range(f"{final}:{final}").delete()
#                 # input_sht.range(f"{final}:{final}").color ="#00FF00"
#                 # if not len(row_dict["Knock_Off"]):
#                 #     row_dict["Knock_Off"] = [[f"{curr}:{final}"]]                   
#                 # elif len(row_dict["Knock_Off"][-1]) <=24:
#                 #     if int(row_dict["Knock_Off"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1                       
#                 #         row_dict["Knock_Off"][-1][-1] = f'{row_dict["Knock_Off"][-1][-1].split(":")[0]}:{final}'
#                 #     else:
#                 #         row_dict["Knock_Off"][-1].append(f"{curr}:{final}")
#                 # elif len(row_dict["Knock_Off"][-1]) >24:
#                     # row_dict["Knock_Off"].append([f"{curr}:{final}"])
#             # curr-=1
#         elif (abs(input_sht.range(f"{credit_col_letter}{curr}").value) - abs(input_sht2.range(f"{debit_col_letter}{final}").value))<10:
#             #amt diff
#             print(f"Moving {curr} to amount diff")
#             amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row

#             if input_sht==input_sht2:
#                 pass
#                 # input_sht.range(f"{curr}:{final}").api.Copy()
#                 # wb.activate()
#                 # amt_diff_sht.activate()
#                 # amt_diff_sht.range(f"A{amt_diff_last_row+1}").api.Select()
#                 # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                 # amt_diff_sht.autofit()
#                 # input_sht.range(f"{i}:{i+1}").copy(amt_diff_sht.range(f"A{amt_diff_last_row+1}"))

#                 # input_sht.range(f"{curr}:{final}").delete()
#                 # input_sht.range(f"{curr}:{final}").color ="#FFFF00"
#                 if not len(row_dict["Amt_Dff"]):
#                     row_dict["Amt_Dff"] = [[f"{curr}:{final}"]]                   
#                 elif len(row_dict["Amt_Dff"][-1]) <=24:
#                     if int(row_dict["Amt_Dff"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1                       
#                         row_dict["Amt_Dff"][-1][-1] = f'{row_dict["Amt_Dff"][-1][-1].split(":")[0]}:{final}'
#                     else:
#                         row_dict["Amt_Dff"][-1].append(f"{curr}:{final}")
#                 elif len(row_dict["Amt_Dff"][-1]) >24:
#                     row_dict["Amt_Dff"].append([f"{curr}:{final}"])
#             else:
#                 input_sht.range(f"{curr}:{curr}").api.Copy()
#                 wb.activate()
#                 amt_diff_sht.activate()
#                 amt_diff_sht.range(f"A{amt_diff_last_row+1}").api.Select()
#                 wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                
#                 input_sht2.range(f"B{final}:{eth_trueup_col_letter}{final}").api.Copy()
#                 wb.activate()
#                 amt_diff_sht.activate()
#                 amt_diff_sht.range(f"A{amt_diff_last_row+2}").api.Select()
#                 wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

#                 amt_diff_sht.range(f"K{knock_off_last_row+2}").copy(amt_diff_sht.range(f"L{knock_off_last_row+2}"))
#                 amt_diff_sht.range(f"K{knock_off_last_row+2}").clear()
#                 amt_diff_sht.range(f"M{knock_off_last_row+2}").clear()#Clearing Final Amount
                

#                 amt_diff_sht.autofit()
#                 # input_sht.range(f"{i}:{i+1}").copy(amt_diff_sht.range(f"A{amt_diff_last_row+1}"))

#                 input_sht.range(f"{curr}:{curr}").delete()
#                 # input_sht.range(f"{curr}:{curr}").color ="#FFFF00"
#                 input_sht2.range(f"{final}:{final}").delete()
#                 # input_sht.range(f"{final}:{final}").color ="#FFFF00"
#                 curr-=1

#                 if not len(row_dict["Amt_Dff"]):
#                     row_dict["Amt_Dff"] = [[f"{curr}:{final}"]]                   
#                 elif len(row_dict["Amt_Dff"][-1]) <=24:
#                     if int(row_dict["Amt_Dff"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1                       
#                         row_dict["Amt_Dff"][-1][-1] = f'{row_dict["Amt_Dff"][-1][-1].split(":")[0]}:{final}'
#                     else:
#                         row_dict["Amt_Dff"][-1].append(f"{curr}:{final}")
#                 elif len(row_dict["Amt_Dff"][-1]) >24:
#                     row_dict["Amt_Dff"].append([f"{curr}:{final}"])

#             # curr-=1
#         else:
#             #line for ethnaol accrual tab
#             print(f'current line {curr} remains here for ethanol accrual tab having mrn no.{input_sht.range(f"{mrn_col_letter}{curr}")}')
#         return curr, row_dict
#     except Exception as e:
#         raise e


# def row_range_calc(filter_col:str, input_sht,wb):
#     sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row

#     sp_address= input_sht.api.Range(f"{filter_col}2:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address

#     sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

#     row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

#     while row_range[-1]!=sp_lst_row:

#         sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row

#         sp_address.extend(input_sht.api.Range(f"{filter_col}{row_range[-1]}:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)

#         # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

#         # row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
        
    
#     sp_address = sp_address.replace("$","").split(",")
#     init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
#     sublist = []
#     flat_list = [item for sublist in init_list for item in sublist]
#     return flat_list, sp_lst_row,sp_address

# def thick_bottom_border(cellrange:str,working_sheet,working_workbook):
#         working_sheet.api.Range(cellrange).Select()
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeLeft).LineStyle = win32c.Constants.xlNone
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeTop).LineStyle = win32c.Constants.xlNone
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeRight).LineStyle = win32c.Constants.xlNone
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideVertical).LineStyle = win32c.Constants.xlNone
#         working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideHorizontal).LineStyle = win32c.Constants.xlNone
#         linestylevalues=[win32c.BordersIndex.xlEdgeBottom]
#         for values in linestylevalues:
#             a=working_workbook.app.selection.api.Borders(values)
#             a.LineStyle = win32c.LineStyle.xlContinuous
#             a.ColorIndex = 0
#             a.TintAndShade = 0
#             a.Weight = win32c.BorderWeight.xlMedium

def open_gr(input_date,output_date):
    try:
        msg = openGr(input_date, output_date)
        return msg
    except Exception as e:
        raise e

def unbilled_ar(input_date, output_date):
    try:
        msg = unbilled_ar(input_date, output_date)
        return msg
    except Exception as e:
        raise e

def purchased_ar(input_date, output_date):
    try:
        msg = purchased_ar(input_date, output_date)
        return msg
    except Exception as e:
        raise e


def ar_ageing_bulk(input_date, output_date):
    try:
        msg = ar_ageing_bulk(input_date, output_date)
        return msg
    except Exception as e:
        raise e

def ar_ageing_rack(input_date, output_date):
    try:
        msg = ar_ageing_rack(input_date, output_date)
        return msg
    except Exception as e:
        raise e


def bbr_nlv_futures(start_date,end_date):
    try:
        msg = NLV_FUTURES(start_date,end_date)
        return msg
    except Exception as e:
        raise e

def bbr_cash(start_date,end_date):
    try:
        msg = cash(start_date,end_date)
        return msg
    except Exception as e:
        raise e

# def openGr(input_date, output_date):
#     try:
#         start_time = datetime.now()
#         input_datetime = datetime.strptime(input_date, "%m.%d.%Y")
#         month = input_datetime.month
#         day = input_datetime.day
#         j_loc = r"J:\India\BBR\IT_BBR\Reports"
#         # curr_loc = os.getcwd()
#         # input_sheet= curr_loc+r'\Raw Files'+f'\\Open GR {month}{day}.xlsx'
#         input_sheet= j_loc+r'\Open GR\Raw Files'+f'\\Open GR {month}{day}.xlsx'
#         output_location = j_loc+r'\Open GR\Output Files' 
#         # output_location = curr_loc+r'\Output Files' 
#         if not os.path.exists(input_sheet):
#             return(f"{input_sheet} Excel file not present for date {input_date}")                 
#         retry=0
#         while retry < 10:
#             try:
#                 wb = xw.Book(input_sheet,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e
#         #make copy of Sheet1
#         wb.sheets["Input"].copy(name="Input_Main", after=wb.sheets["Input"])
#         input_sht = wb.sheets["Input_Main"]
        
       
#         #Deleting extras
#         input_sht.range("A:A").api.Delete()
#         input_sht.range(f'1:{input_sht.range("A1").end("down").end("down").row-1}').api.Delete()

#         #Checking Opening Balance
#         curr_col_list = input_sht.range("A1").expand('right').value
#         balance_row = input_sht.range(f'{num_to_col_letters(len(curr_col_list))}1').end('down').row -1
#         balance = input_sht.range(f"{num_to_col_letters(curr_col_list.index('Balance')+1)}{balance_row}").value

#         reco_sht = wb.sheets["Reco"]
#         reco_last_row = reco_sht.range(f'A'+ str(reco_sht.cells.last_cell.row)).end('up').row
#         reco_col_list = reco_sht.range("A1").expand('right').value
#         reco_a_list = reco_sht.range(f"A1:A{reco_last_row}").value
#         input_total = wb.sheets["Input"].range(f"AB{wb.sheets['Input'].range(f'AC'+ str(wb.sheets['Input'].cells.last_cell.row)).end('up').row}").address
#         #Updating reco input sheet value
#         reco_sht.range("B8").formula = f"=+'Input'!{input_total}"
#         # if balance != reco_sht.range(f'{num_to_col_letters(len(reco_col_list))}{reco_a_list.index("Open MRN as Per BS")+1}').value:
#         #     return "Opening blanace of Input Sheet not balanced with Reco sheet"


        







#         #Extra Column deletion logic
#         req_col_list = ["Date", "Cost Center", "Terminal", "Voucher No", "Name", "Vendor Ref", "Pur VNo", "MRN No:", "BOL Number", "Rail Car/Truck #",
#         "Narration"	"Remarks", "Debit Amount", "Credit Amount", "Balance"]
        
#         i=0
#         while len(req_col_list) <=len(curr_col_list):
#             curr_col = num_to_col_letters(i+1)
#             if input_sht.range(f"{curr_col}1").value not in req_col_list:
#                 input_sht.range(f"{curr_col}:{curr_col}").api.Delete()
#                 i-=1
#             curr_col_list = input_sht.range("A1").expand('right').value
#             i+=1
#         #Delete extra rows of total in starting
#         to_be_deleted = input_sht.range("A1").end('down').row
#         input_sht.range(f"2:{to_be_deleted-1}").api.Delete()
#         #Sorting by railcar
#         curr_last_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
#         curr_last_col = len(curr_col_list)
#         curr_last_col_letter = num_to_col_letters(curr_last_col)
#         railcar_col = curr_col_list.index("Rail Car/Truck #")
#         railcar_col_letter = num_to_col_letters(railcar_col+1)
        

#         input_sht.range(f"A1:{curr_last_col_letter}{curr_last_row}").api.Sort(Key1=input_sht.range(f"{railcar_col_letter}1:{railcar_col_letter}{curr_last_row}").api,
#             Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

#         # #Removing Extra total
#         to_be_deleted_final = input_sht.range(f'B'+ str(input_sht.cells.last_cell.row)).end('up').row
#         to_be_deleted_init = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row

#         input_sht.range(f"{to_be_deleted_init+1}:{to_be_deleted_final+5}").api.Delete()#+% for deleting extra line border
#         input_sht.copy(name = "Input_Main2", after = wb.sheets["Input_Main"])
#         input_sht = wb.sheets["Input_Main2"]

#         voucher_col = curr_col_list.index("Voucher No")
#         voucher_col_col_letter = num_to_col_letters(voucher_col+1)
        
#         mrn_col = curr_col_list.index("MRN No:")
#         mrn_col_letter = num_to_col_letters(mrn_col+1)

#         date_col = curr_col_list.index("Date")
#         date_col_letter = num_to_col_letters(date_col+1)


#         debit_col = curr_col_list.index("Debit Amount")
#         debit_col_letter = num_to_col_letters(debit_col+1)

#         credit_col = curr_col_list.index("Credit Amount")
#         credit_col_letter = num_to_col_letters(credit_col+1)

#         bol_col = curr_col_list.index("BOL Number")
#         bol_col_letter = num_to_col_letters(bol_col+1)


#         last_row = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
#         curr_month_num =datetime.strptime(input_date,"%m.%d.%Y").month
#         curr_month = datetime.strftime(datetime.strptime(input_date,"%m.%d.%Y"), "%b")
#         prev_month = datetime.strftime((datetime.strptime(input_date,"%m.%d.%Y").replace(day=1) -timedelta(days=1)), "%b")


#         #Adding all sheets at once #Logic avoided as these sheets presaent from previous file
#         # knock_off_sht = wb.sheets.add("Knocked Off",after=wb.sheets[-1])
#         # amt_diff_sht = wb.sheets.add("Amount Diff",after=wb.sheets[-1])
#         # diff_month_sht = wb.sheets.add(f"{prev_month} MRN booked in {curr_month}",after=wb.sheets[-1])

#         knock_off_sht = wb.sheets("Knocked Off")
#         amt_diff_sht = wb.sheets("Amount Diff")
#         try:
#             diff_month_sht = wb.sheets(f"{prev_month} MRN booked in {curr_month}")
#         except:
#             diff_month_sht = wb.sheets.add(f"{prev_month} MRN booked in {curr_month}",after=amt_diff_sht)

#         #Adding headers in all new sheets
#         input_sht.range(f"A1").expand("right").copy(knock_off_sht.range("A1"))
#         input_sht.range(f"A1").expand("right").copy(amt_diff_sht.range("A1"))
#         input_sht.range(f"A1").expand("right").copy(diff_month_sht.range("A1"))
#         ignore_check= False
#         if day == 15:#replace else append

#             knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
#             knock_off_sht.range(f"A2:A{knock_off_last_row}").api.EntireRow.Delete()
#             amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
#             amt_diff_sht.range(f"A2:A{amt_diff_last_row}").api.EntireRow.Delete()
#             diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row
#             if diff_month_last_row!=1:
#                 diff_month_sht.range(f"A2:A{diff_month_last_row}").api.EntireRow.Delete()

#         knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
#         amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
#         diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row

#         i=2
#         row_dict = {}
#         row_dict["Knock_Off"] = []
#         row_dict["Amt_Dff"] = []
#         # amtdiff_dict = {}
        
#         while i <=last_row:
#             if not ignore_check:
#                 #Checking Mrn with next pjv row
#                 if input_sht.range(f"{voucher_col_col_letter}{i}").value.split(":")[1] == input_sht.range(f"{mrn_col_letter}{i+1}").value:
#                     #Condition for knock off and amount diff tab
                    
#                     if input_sht.range(f"{date_col_letter}{i}").value.month == curr_month_num:
#                         #knock Off
#                         if input_sht.range(f"{credit_col_letter}{i}").value is not None and input_sht.range(f"{debit_col_letter}{i+1}").value is not None:
#                             i, row_dict = knockOffAmtDiff(i, i+1, wb, input_sht, input_sht, credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter, row_dict)
#                         else:#interchange debit and credit col
#                             i, row_dict = knockOffAmtDiff(i, i+1, wb, input_sht, input_sht, debit_col_letter, credit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter, row_dict)




                        
#                         last_row = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
#                         # ignore_check=True
#                         # print("Move both enteries to knock off tab")
#                     #prev month MRN Booked in Current Month
#                     elif input_sht.range(f"{date_col_letter}{2}").value.month != curr_month_num:
#                         print("Move both enteries to prev month MRN Booked in Current Month")
#                         diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row

#                         input_sht.range(f"{i}:{i+1}").api.Copy()
#                         wb.activate()
#                         diff_month_sht.activate()
#                         diff_month_sht.range(f"A{diff_month_last_row+1}").api.Select()
#                         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                         diff_month_sht.autofit()

#                         # input_sht.range(f"{i}:{i+1}").copy(diff_month_sht.range(f"A{diff_month_last_row+1}"))

#                         input_sht.range(f"{i}:{i+1}").api.Delete()

#                         i-=1
#                     else:
#                         print(f"New case for row number {i}")
#                 else:
#                     print(f"MRN no or pjv line not found in row {i}",end="\n")
#                     print(f"Keeping this row for ethanol accrual")
#             else:
#                 print(f"pjv row num is {i}")
#             i+=1
        
#         ###########################Copy pasting based on lista###################################################################
#         colorList = []
#         for key in row_dict.keys():
    
#             for rowList in row_dict[key]:
#                 rows = ",".join(rowList)
#                 if key == "Knock_Off":
#                     knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
#                     input_sht.range(rows).copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))
#                     input_sht.range(rows).color = "#00FF0"
#                     if input_sht.range(rows).api.Interior.Color not in colorList:
#                         colorList.append(input_sht.range(rows).api.Interior.Color)
#                 else:
#                     wb.activate()
#                     input_sht.activate()
#                     input_last_row1 = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row +3
#                     input_sht.range(rows).copy(input_sht.range(f"A{input_last_row1}"))
#                     input_last_row2 = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
#                     amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
#                     input_sht.range(f"{input_last_row1}:{input_last_row2}").api.Copy()

#                     wb.activate()
#                     amt_diff_sht.activate()
#                     amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
#                     amt_diff_sht.range(f"A{amt_diff_last_row+1}").api.Select()
#                     wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                     amt_diff_sht.autofit()
#                     input_sht.activate()
#                     input_sht.range(rows).color = "#FFFF00"
#                     if input_sht.range(rows).api.Interior.Color not in colorList:
#                         colorList.append(input_sht.range(rows).api.Interior.Color)

#                     input_sht.range(f"{input_last_row1}:{input_last_row2}").delete()

#         ###########################Deletion Logic#################################################################################
#         for colors in colorList:
#             input_sht.activate()
#             input_sht.api.AutoFilterMode=False
#             input_sht.api.Range(f"{railcar_col_letter}1").AutoFilter(Field:=f"{railcar_col+1}", Criteria1:=colors, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#             fil_last_row = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
#             if fil_last_row !=1:
#                 input_sht.range(f"2:{fil_last_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         ##########################################################################################################################


#         input_sht.api.AutoFilterMode=False
#         #Filtering out remaining
#         input_sht.autofit()
#         input_sht.activate()
#         font_colour,Interior_colour = conditional_formatting(f"{railcar_col_letter}:{railcar_col_letter}",input_sht,wb)
#         input_sht.api.AutoFilterMode=False
#         input_sht.api.Range(f"{railcar_col_letter}1").AutoFilter(Field:=f"{railcar_col+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#         input_sht.range(f"A1:{curr_last_col_letter}{curr_last_row}").api.Sort(Key1=input_sht.range(f"{railcar_col_letter}1:{railcar_col_letter}{curr_last_row}").api,
#             Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

#         #Finding filtered range

#         row_range, sp_lst_row, sp_address = row_range_calc('A', input_sht, wb)
#         curr_railcar = input_sht.range(f"{railcar_col_letter}{row_range[0]}").value
#         curr_index = 0
#         final_index = 0
#         i=0
#         # for i in row_range:
#         while sp_lst_row!=1:
#             if (input_sht.range(f"{railcar_col_letter}{row_range[i]}").value!=curr_railcar) or (row_range[i] == row_range[-1]):
#                 input_sht.activate()
#                 final_index = i-1
                
#                 if row_range[i] == row_range[-1]:
#                     final_index = i
                    
#                 #sum of Debit amount and Credit Amount
#                 debit_value = input_sht.range(f"{debit_col_letter}{row_range[curr_index]}:{debit_col_letter}{row_range[final_index]}").value
#                 credit_value = input_sht.range(f"{credit_col_letter}{row_range[curr_index]}:{credit_col_letter}{row_range[final_index]}").value
#                 if isinstance(debit_value, list):
#                     debit_sum = sum(filter(None, debit_value))
#                 else:
#                     debit_sum = debit_value
#                 if isinstance(credit_value, list):
#                     credit_sum = sum(filter(None, credit_value))
#                 else:
#                     credit_sum = credit_value
#                 if (debit_sum+credit_sum) == 0:
#                     # if input_sht.range(f"{credit_col_letter}{row_range[curr_index]}").value is not None and input_sht.range(f"{debit_col_letter}{row_range[final_index]}").value is not None:
#                     knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
#                     input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))
#                     input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").api.Delete()

                    
#                     #     i = knockOffAmtDiff(row_range[curr_index], row_range[final_index], wb, input_sht, credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)
#                     # else:#interchange debit and credit col
#                     #     i = knockOffAmtDiff(row_range[curr_index], row_range[final_index], wb, input_sht, debit_col_letter, credit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)
#                     # i = knockOffAmtDiff(row_range[curr_index], row_range[final_index], wb, input_sht, credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)
#                     row_range, sp_lst_row, sp_address = row_range_calc('A', input_sht, wb)
#                     curr_railcar = input_sht.range(f"{railcar_col_letter}{row_range[0]}").value
#                     curr_index = 0
#                     i=0
#                 else:
#                     print("New condition found moving that data to Special_Sheet")
#                     try:
#                         spcl_sht = wb.sheets["Special_Sheet"]
#                     except:
#                         spcl_sht = wb.sheets.add(name="Special_Sheet", after=reco_sht)

#                     input_sht.range(f"A1").expand("right").copy(spcl_sht.range("A1"))
#                     spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row


#                     input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").copy(spcl_sht.range(f"A{spcl_sht_last_row+1}"))

#                     input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").api.Delete()
#                     row_range, sp_lst_row, sp_address = row_range_calc('A', input_sht, wb)
#                     curr_railcar = input_sht.range(f"{railcar_col_letter}{row_range[0]}").value
#                     curr_index = 0
#                     i=0

            
#                 # curr_index=final_index
#                 i-=1
            
            
            
            
#             sp_lst_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
#             i+=1

#         #################################Add logic again copy back data from special sheet to input sheet#########################        
#         input_sht.api.AutoFilterMode=False
#         spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row
#         last_row = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
#         spcl_sht.range(f"2:{spcl_sht_last_row}").copy(input_sht.range(f"A{last_row+1}"))

#         #Deleting copied data from special sheet
#         spcl_sht.range(f"2:{spcl_sht_last_row}").api.Delete()

#         input_sht.range(f"A1:{curr_last_col_letter}{curr_last_row}").api.Sort(Key1=input_sht.range(f"{railcar_col_letter}1:{railcar_col_letter}{curr_last_row}").api,
#             Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

#         #MRR will be donw at end
#         #Now pjv logic

        
#         input_sht.api.Range(f"{voucher_col_col_letter}1").AutoFilter(Field:=f"{voucher_col+1}", Criteria1:="Pjv*", Operator:=7)
#         sp_lst_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
#         try:
#             pjv_sht = wb.sheets.add("PJV",after=input_sht)
#         except:
#             pjv_sht = wb.sheets("PJV")
#         input_sht.activate()
#         input_sht.api.Range(f"A1:{curr_last_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         wb.app.selection.copy(pjv_sht.range(f"A1"))
#         pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row

#         input_sht.activate()
#         input_sht.api.Range(f"A2:{curr_last_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         input_sht.api.AutoFilterMode=False
        

#         #Add MRN move to logic here

#         ###Ethanol Accrual Sheet logic starts here
#         eth_acr_sht = wb.sheets("Ethanol Accrual")
#         eth_col_list = eth_acr_sht.range("A1").expand('right').value
#         eth_credit_col = eth_col_list.index("Credit Amount")
#         eth_credit_col_letter = num_to_col_letters(eth_credit_col+1)

#         eth_final_amt_col = eth_col_list.index("Final Amount")
#         eth_final_amt_col_letter = num_to_col_letters(eth_final_amt_col+1)

#         eth_rail_col = eth_col_list.index("Rail Car/Truck #")
#         eth_rail_col_letter = num_to_col_letters(eth_rail_col+1)

#         eth_last_col = len(eth_col_list)
#         eth_last_col_letter = num_to_col_letters(eth_last_col)
#         eth_trueup_col = eth_col_list.index("TrueUp")
#         eth_trueup_col_letter = num_to_col_letters(eth_trueup_col+1)
#         eth_bol_col = eth_col_list.index("BOL Number")
#         eth_bol_col_letter = num_to_col_letters(eth_bol_col+1)

#         #filter out red color cell from credit amount column
#         eth_acr_sht.api.AutoFilterMode=False
#         eth_acr_sht.api.Range(f"{eth_credit_col_letter}1").AutoFilter(Field:=f"{eth_credit_col+1}", Criteria1:=Interior_colour, 
#         Operator:=win32c.AutoFilterOperator.xlFilterNoFill)
#         eth_acr_sht.activate()
#         sp_lst_row = eth_acr_sht.range(f'A'+ str(eth_acr_sht.cells.last_cell.row)).end('up').row
        
        
        
#         pjv_col_list = pjv_sht.range(f"A1").expand('right').value
#         pjv_last_col = len(pjv_col_list)
#         pjv_last_col_letter = num_to_col_letters(pjv_last_col+1)

#         pjv_trueup_col = pjv_last_col+1
#         pjv_trueup_col_letter = num_to_col_letters(pjv_trueup_col+1)

#         pjv_credit_col = pjv_col_list.index("Credit Amount")
#         pjv_credit_col_letter = num_to_col_letters(pjv_credit_col+1)

#         pjv_debit_col = pjv_col_list.index("Debit Amount")
#         pjv_debit_col_letter = num_to_col_letters(pjv_debit_col+1)

#         pjv_mrn_col = pjv_col_list.index("MRN No:")
#         pjv_mrn_col_letter = num_to_col_letters(pjv_mrn_col+1)
        
#         pjv_railcar_col = pjv_col_list.index("Rail Car/Truck #")
#         pjv_railcar_col_letter = num_to_col_letters(pjv_railcar_col+1)

#         pjv_bol_col = pjv_col_list.index("BOL Number")
#         pjv_bol_col_letter = num_to_col_letters(pjv_bol_col+1)

#         pjv_voucher_col = pjv_col_list.index("Voucher No")
#         pjv_voucher_col_letter = num_to_col_letters(pjv_voucher_col+1)

#         pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row
        
        
#         #Pasting BOL numbers from pjv sheet
#         # pjv_sht.range(f"{pjv_bol_col_letter}2:{pjv_bol_col_letter}{pjv_last_row}").copy(eth_acr_sht.range(f"{eth_bol_col_letter}{sp_lst_row+6}"))
#         #using railcar instead of bol number for getting data from ethanol accrual sheet
#         pjv_sht.range(f"{pjv_railcar_col_letter}2:{pjv_railcar_col_letter}{pjv_last_row}").copy(eth_acr_sht.range(f"{eth_rail_col_letter}{sp_lst_row+6}"))

#         font_colour,Interior_colour = conditional_formatting(f"{eth_rail_col_letter}:{eth_rail_col_letter}",eth_acr_sht,wb)
#         # eth_acr_sht.api.AutoFilterMode=False
#         eth_acr_sht.api.Range(f"{eth_rail_col_letter}1").AutoFilter(Field:=f"{eth_rail_col+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        
        
#         eth_acr_sht.api.Range(f"B1:{eth_trueup_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         wb.app.selection.copy(pjv_sht.range(f"A{pjv_last_row+1}"))

#         #deleting bol numbers copied from pjv sheet in eth accr sheet
#         eth_acr_sht.range(f"{eth_bol_col_letter}{sp_lst_row+6}").expand("down").clear()


#         #Deleting copied data from ethanol Accrual Sheet
#         eth_acr_sht.api.Range(f"A2:{eth_last_col_letter}{sp_lst_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()#Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         wb.app.selection.delete(shift='left')
#         eth_acr_sht.api.Range(f"A2:{eth_last_col_letter}{sp_lst_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()#Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         wb.app.selection.delete(shift='up')
#         # input_sht.api.AutoFilterMode=False

# ############################################Update logic from above for adding bol number of pjv for duplicate check #########################################################################################################################################
#         pjv_sht.activate()
#         pjv_col2_list = pjv_sht.range(f"A{pjv_last_row+1}").expand('right').value
#         pjv_fin_amt2_col = pjv_col2_list.index("Final Amount")
#         pjv_fin_amt2_col_letter = num_to_col_letters(pjv_fin_amt2_col+1)
#         pjv_trueup2_col = pjv_col2_list.index("TrueUp")
#         pjv_trueup2_col_letter = num_to_col_letters(pjv_trueup2_col+1)
#         pjv_credit2_col = pjv_col2_list.index("Credit Amount")
#         pjv_credit2_col_letter = num_to_col_letters(pjv_credit2_col+1)

        

        
        

#         # #HighLighting duplicate Railcar numbers
#         # font_colour,Interior_colour = conditional_formatting(pjv_railcar_col,pjv_sht,wb)

#         # pjv_sht.api.AutoFilterMode=False
#         # pjv_sht.api.Range(f"{pjv_railcar_col_letter}1").AutoFilter(Field:=f"{pjv_railcar_col+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterNoFill)

#         # pjv_sht.activate()
#         # sp_lst_row = pjv_sht.range(f'A'+ str(pjv_sht.cells.last_cell.row)).end('up').row
#         # pjv_sht.api.Range(f"A2:{pjv_last_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         # wb.app.selection.copy(eth_acr_sht.range(f"A{pjv_last_row+1}"))
#         pjv_sht.activate()
        
#         #Making Trueup Col
#         pjv_sht.range(f"{pjv_trueup_col_letter}1").value = "TrueUp"

#         #Deletion and column shifting logic
#         pjv_sht.range(f"{pjv_fin_amt2_col_letter}{pjv_last_row+1}").expand("down").api.Delete()
#         pjv_col2_list = pjv_sht.range(f"A{pjv_last_row+1}").expand('right').value
#         pjv_trueup2_col = pjv_col2_list.index("TrueUp")
#         pjv_trueup2_col_letter = num_to_col_letters(pjv_trueup2_col+1)
        

#         pjv_sht.range(f"{pjv_trueup2_col_letter}{pjv_last_row+1}").expand("down").api.Cut(pjv_sht.range(f"{pjv_trueup_col_letter}{pjv_last_row+1}").api)
        
#         pjv_sht.range(f"{pjv_credit2_col_letter}{pjv_last_row+1}").expand("down").api.Cut(pjv_sht.range(f"{pjv_credit_col_letter}{pjv_last_row+1}").api)

#         #Deleting secondary headers
#         pjv_sht.range(f"{pjv_last_row+1}:{pjv_last_row+1}").api.Delete()

#         #Sorting by railcar
#         pjv_last_row = pjv_sht.range(f'A'+ str(pjv_sht.cells.last_cell.row)).end('up').row
        
#         pjv_sht.range(f"A1:{pjv_trueup_col_letter}{pjv_last_row}").api.Sort(Key1=pjv_sht.range(f"{pjv_voucher_col_letter}1:{pjv_voucher_col_letter}{pjv_last_row}").api,
#         Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)


#         pjv_sht.range(f"A1:{pjv_trueup_col_letter}{pjv_last_row}").api.Sort(Key1=pjv_sht.range(f"{pjv_railcar_col_letter}1:{pjv_railcar_col_letter}{pjv_last_row}").api,
#         Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

#         row_dict = {}
#         row_dict["Knock_Off"] = []
#         row_dict["Amt_Dff"] = []
#         i=2
#         while i <=pjv_last_row:
#             if not ignore_check:
#                 #Checking Mrn with next pjv row
#                 if pjv_sht.range(f"{voucher_col_col_letter}{i}").value.split(":")[1] == pjv_sht.range(f"{mrn_col_letter}{i+1}").value:
#                     #Condition for knock off and amount diff tab
                    
#                     if pjv_sht.range(f"{date_col_letter}{i}").value.month == curr_month_num:
#                         #knock Off
#                         if pjv_sht.range(f"{pjv_credit_col_letter}{i}").value is not None and pjv_sht.range(f"{debit_col_letter}{i+1}").value is not None:
#                             i, row_dict = knockOffAmtDiff(i, i+1, wb, pjv_sht, pjv_sht, pjv_credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, pjv_mrn_col_letter, row_dict)
#                         else:#interchange debit and credit col
#                             i, row_dict = knockOffAmtDiff(i, i+1, wb, pjv_sht, pjv_sht, pjv_debit_col_letter, pjv_credit_col_letter, knock_off_sht, amt_diff_sht, pjv_mrn_col_letter, row_dict)




                        
#                         pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row
#                         # ignore_check=True
#                         # print("Move both enteries to knock off tab")
#                     #prev month MRN Booked in Current Month
#                     elif pjv_sht.range(f"{date_col_letter}{2}").value.month != curr_month_num:
#                         print("Move both enteries to prev month MRN Booked in Current Month")
#                         diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row

#                         pjv_sht.range(f"{i}:{i+1}").api.Copy()
#                         wb.activate()
#                         diff_month_sht.activate()
#                         diff_month_sht.range(f"A{diff_month_last_row+1}").api.Select()
#                         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                         diff_month_sht.autofit()

#                         # pjv_sht.range(f"{i}:{i+1}").copy(diff_month_sht.range(f"A{diff_month_last_row+1}"))

#                         pjv_sht.range(f"{i}:{i+1}").api.Delete()

#                         i-=1
#                     else:
#                         print(f"New case for row number {i}")
#                 else:
#                     print(f"MRN no or pjv line not found in row {i}",end="\n")
#                     print(f"Keeping this row for ethanol accrual")
#             else:
#                 print(f"pjv row num is {i}")
#             i+=1
#             pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row

#         ###########################Copy pasting based on lista###################################################################
#         colorList = []
#         for key in row_dict.keys():
    
#             for rowList in row_dict[key]:
#                 rows = ",".join(rowList)
#                 if key == "Knock_Off":
#                     knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
#                     pjv_sht.range(rows).copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))
#                     pjv_sht.range(rows).color = "#00FF0"
#                     if pjv_sht.range(rows).api.Interior.Color not in colorList:
#                         colorList.append(pjv_sht.range(rows).api.Interior.Color)
#                 else:
#                     wb.activate()
#                     input_sht.activate()
#                     input_last_row1 = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row +3
#                     input_sht.range(rows).copy(input_sht.range(f"A{input_last_row1}"))
#                     input_last_row2 = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
#                     amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
#                     input_sht.range(f"{input_last_row1}:{input_last_row2}").api.Copy()

#                     wb.activate()
#                     amt_diff_sht.activate()
#                     amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
#                     amt_diff_sht.range(f"A{amt_diff_last_row+1}").api.Select()
#                     wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                     amt_diff_sht.autofit()
#                     input_sht.activate()
#                     input_sht.range(rows).color = "#FFFF00"
#                     if input_sht.range(rows).api.Interior.Color not in colorList:
#                         colorList.append(input_sht.range(rows).api.Interior.Color)

#                     input_sht.range(f"{input_last_row1}:{input_last_row2}").delete()

#         ###########################Deletion Logic#################################################################################
#         for colors in colorList:
#             pjv_sht.activate()
#             pjv_sht.api.AutoFilterMode=False
#             pjv_sht.api.Range(f"{railcar_col_letter}1").AutoFilter(Field:=f"{railcar_col+1}", Criteria1:=colors, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#             fil_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row
#             if fil_last_row !=1:
#                 pjv_sht.range(f"2:{fil_last_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         ##########################################################################################################################

        
#         input_sht.api.AutoFilterMode=False
#         pjv_sht.range(f"A1").expand("right").copy(spcl_sht.range("A1"))
#         try:
#             spcl_sht = wb.sheets["Special_Sheet"]
#         except:
#             spcl_sht = wb.sheets.add(name="Special_Sheet", after=reco_sht)
#         spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row


#         pjv_sht.range(f"2:{pjv_last_row}").copy(spcl_sht.range(f"A{spcl_sht_last_row+1}"))

#         pjv_sht.range(f"2:{pjv_last_row}").api.Delete()


#         #Now deleting pjv Sheet
#         pjv_sht.delete()

#         #Now checking input sheet for remaing rows
#         input_sht.activate()
#         #Removing MRR Logic
#         input_sht.api.AutoFilterMode=False
#         input_sht.api.Range(f"{voucher_col_col_letter}1").AutoFilter(Field:=f"{voucher_col+1}", Criteria1:="MRR*", Operator:=7)

#         #searching all bol numbers in ethanol accrual sheet for each mrr found in inpurt sheet
#         row_range, sp_lst_row, sp_address = row_range_calc('A', input_sht, wb)
#         curr=0
#         for i in range(len(row_range)):

#             bol_num = input_sht.range(f"{bol_col_letter}{row_range[i]}").value
#             eth_acr_sht.activate()
#             eth_acr_sht.api.AutoFilterMode=False
#             try:
#                 eth_acr_sht.api.Cells.Find(What:=bol_num , After:=eth_acr_sht.api.Application.ActiveCell, LookIn:=win32c.FindLookIn.xlFormulas,
#                 LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()

#                 cell_value = eth_acr_sht.api.Application.ActiveCell.Address.replace("$","")
#                 row_num = int(re.findall(r'\d+', cell_value)[0])

#                 #Copy delete logic
#                 curr=knockOffAmtDiff(row_range[i],row_num, wb, input_sht, eth_acr_sht, debit_col_letter, eth_credit_col_letter, knock_off_sht, amt_diff_sht,
#                 mrn_col_letter, eth_trueup_col_letter)
#                 curr = row_range[i]-curr
                

#             except:
#                 spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row
#                 input_sht.range(f"{row_range[i]-curr}:{row_range[i]-curr}").copy(spcl_sht.range(f"A{spcl_sht_last_row+1}"))
            
#                 input_sht.range(f"{row_range[i]-curr}:{row_range[i]-curr}").api.Delete()


#         #Logic for moving remaining mrn in input sheet to ethanol accrual sheet
#         input_sht.api.AutoFilterMode=False
#         curr_last_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
#         eth_last_row = eth_acr_sht.range(f'A'+ str(eth_acr_sht.cells.last_cell.row)).end('up').row

#         input_sht.activate()
#         row_count = input_sht.range(f"A2").expand("down").count
#         for i in range(0,row_count):
#             eth_acr_sht.api.Range(f"B{eth_last_row+1}").EntireRow.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#         input_sht.range(f"A2:{credit_col_letter}{curr_last_row}").copy(eth_acr_sht.range(f"B{eth_last_row+1}"))
#         input_sht.range(f"A2:{credit_col_letter}{curr_last_row}").api.EntireRow.Delete()
#         eth_acr_sht.range(f"M{eth_last_row+1}").expand("down").copy(eth_acr_sht.range(f"L{eth_last_row+1}"))
#         # eth_acr_sht.range(f"M{eth_last_row+1}").expand("down").clear()
#         eth_acr_sht.range(f"M{eth_last_row+1}").expand("down").api.NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         eth_acr_sht.range(f"L{eth_last_row+1}").expand("down").api.NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        
#         eth_acr_sht.activate()
#         #Refreshing pivot table in ethanol accrual tab
#         pivotCount = wb.api.ActiveSheet.PivotTables().Count
#          # 'INPUT DATA'!$A$3:$I$86
#         for j in range(1, pivotCount+1):     
#             wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

#         wb.save(output_location+f"\\Open GR {month}{day}.xlsx")
#         end_time = datetime.now()
#         total_time = end_time - start_time
#         print(f"Total time taken {total_time}")

#         print("Done")
#         return(f"Open GR report for {input_date} has been generated successfully")
#     except Exception as e:
#         raise e
#     finally:
#         try:
#             wb.app.quit()
#         except:
#             pass


# def ar_ageing_bulk(input_date, output_date):
#     try:
#         today_date=date.today()     
#         job_name = 'ar_ageing_Bulk'
#         month = input_date.split(".")[0]
#         day = input_date.split(".")[1]
#         year = input_date.split(".")[-1]
#         input_sheet= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Input'+f'\\AR Aging Bulk {month}{day}.xlsx'
#         output_location = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Output'
#         input_sheet2= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Input'+f'\\BS Bulk {month}{day}.xlsx'
#         input_sheet3= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Template_File'+f'\\Biourja_mapping.xlsx'
#         input_sheet4 = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Template_File'+f'\\AR Aging Bulk Template.xlsx'
#         grp_sheet = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Template_File'+f'\\Group_mapping.xlsx'
#         if not os.path.exists(input_sheet):
#             return(f"{input_sheet} Excel file not present for date {input_date}")  
#         if not os.path.exists(input_sheet2):
#             return(f"{input_sheet2} Excel file not present for date {input_date}")  
#         if not os.path.exists(input_sheet3):
#             return(f"{input_sheet3} Excel file not present")    
#         if not os.path.exists(input_sheet4):
#             return(f"{input_sheet4} Excel file not present")                       
#         raw_df = pd.read_excel(input_sheet)    
#         raw_df = raw_df[(raw_df[raw_df.columns[0]] == 'Demurrage')]
#         raw_df = raw_df.iloc[:,[0,1,-6,-5,-4,-3,-2,-1]]
#         raw_df.columns = ['dem_check',"Customer","Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]
#         retry=0
#         while retry < 10:
#             try:
#                 temp_wb = xw.Book(input_sheet4,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e                     
#         retry=0
#         while retry < 10:
#             try:
#                 wb = xw.Book(input_sheet,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         initial_tab= wb.sheets[0]
#         initial_tab.api.Copy(After=wb.api.Sheets(1))
#         input_tab = wb.sheets[1]
        
#         input_tab.name = "Updated_Data(IT)"

#         # check_column= input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         # if check_column ==1:
#         input_tab.api.Range(f"A:A").EntireColumn.Delete()   

#         input_tab.api.Range(f"1:5").EntireRow.Delete()
#         input_tab.api.Range(f"I:L").EntireColumn.Delete() 
#         input_tab.autofit()
#         input_tab.api.Range(f"2:2").EntireRow.Delete()
#         input_tab.activate()


#         column_list = input_tab.range("A1").expand('right').value
#         Voucher_No_column_no = column_list.index('Voucher No')+1
#         Voucher_No_column_letter=num_to_col_letters(Voucher_No_column_no)
#         last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)

#         dict1={"<>":[Voucher_No_column_no,Voucher_No_column_letter,"B"]}
#         for key, value in dict1.items():
#             try:
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
#                 time.sleep(1)
#                 sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                 sp_address= input_tab.api.Range(f"{value[2]}2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                 sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#             except:
#                 pass  

#         input_tab.range(f"Q1").value = "Diff"
#         input_tab.range(f"Q{sp_initial_rw}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         input_tab.range(f"Q{sp_initial_rw}").value=f'=+K{sp_initial_rw}-SUM(L{sp_initial_rw}:P{sp_initial_rw})'
#         lsr_rw = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
#         input_tab.api.Range(f"{lsr_rw+1}:{lsr_rw+10}").EntireRow.Delete()
#         input_tab.api.Range(f"Q{sp_initial_rw}:Q{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         wb.app.api.Selection.FillDown()
#         input_tab.autofit()
#         freezepanes_for_tab(cellrange="2:2",working_sheet=input_tab,working_workbook=wb)


#         for i in range(2,int(f'{lsr_rw}')):
#             if (input_tab.range(f"B{i}").value=="Opb:OPB-1624" or input_tab.range(f"J{i}").value=="Opb:OPB-1624") and int(input_tab.range(f"K{i}").value)==58343:
#                 print(f"deleted customer={input_tab.range(f'A{i}').value} and deleted row={i}")
#                 input_tab.range(f"{i}:{i}").api.Delete()
#                 break
#             else:
#                 pass  

#         input_tab.range(f"Q{sp_initial_rw}:Q{sp_lst_row}")
        
#         voucher_filters = input_tab.range(f"B2:B{sp_lst_row}").value
#         jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Jrn" in filter]
#         input_tab.api.AutoFilterMode=False
#         for value in jeneral_entry:
#             for index,filter in value.items():
#                 try:
#                     input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
#                     time.sleep(1)
#                     sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                     sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                     sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
#                     if messagebox.askyesno("Jrn Entry Found!!!",'Do you want this entry to be removed'):
#                         print("remove entry") 
#                         company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
#                         input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").api.Delete()
#                         input_tab.api.AutoFilterMode=False 
#                         input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
#                         sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#                         sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                         sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
#                         length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
#                         if length <=1:
#                            input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").api.Delete() 
#                            input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").api.Delete()
#                         else:
#                             print("Entries found hence no bucket deletion")
#                         input_tab.api.AutoFilterMode=False
#                     else:
#                         print("continue")
#                         input_tab.range(f"D{index}").copy(input_tab.range(f"E{index}"))
#                         diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{index}").value,"%m-%d-%Y")).days
#                         if diff <11:
#                             input_tab.range(f"K{index}").copy(input_tab.range(f"L{index}"))
#                         elif diff >=11 and diff <31:
#                             input_tab.range(f"K{index}").copy(input_tab.range(f"M{index}"))
#                         elif diff >=31 and diff <61:
#                             input_tab.range(f"K{index}").copy(input_tab.range(f"N{index}"))
#                         elif diff >=61 and diff <91:
#                             input_tab.range(f"K{index}").copy(input_tab.range(f"O{index}"))
#                         else:
#                             input_tab.range(f"K{index}").copy(input_tab.range(f"P{index}"))
#                         input_tab.api.AutoFilterMode=False    
#                 except:
#                     pass   

#         jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Exc" in filter]
#         input_tab.api.AutoFilterMode=False
#         for value in jeneral_entry:
#             for index,filter in value.items():
#                 try:
#                     input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
#                     time.sleep(1)
#                     sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                     sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                     sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
#                     if messagebox.askyesno("Exc Entry Found!!!",'Do you want this entry to be removed'):
#                         print("remove entry") 
#                         company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
#                         input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").api.Delete()
#                         input_tab.api.AutoFilterMode=False 
#                         input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
#                         sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#                         sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                         sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
#                         length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
#                         if length <=1:
#                            input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").api.Delete() 
#                            input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").api.Delete()
#                         else:
#                             print("Entries found hence no bucket deletion")
#                         input_tab.api.AutoFilterMode=False
#                     else:
#                         print("continue")
#                         input_tab.range(f"D{sp_initial_rw_ex}").copy(input_tab.range(f"E{sp_initial_rw_ex}"))
#                         diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{sp_initial_rw_ex}").value,"%m-%d-%Y")).days
#                         if diff <11:
#                             input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"L{sp_initial_rw_ex}"))
#                         elif diff >=11 and diff <31:
#                             input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"M{sp_initial_rw_ex}"))
#                         elif diff >=31 and diff <61:
#                             input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"N{sp_initial_rw_ex}"))
#                         elif diff >=61 and diff <91:
#                             input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"O{sp_initial_rw_ex}"))
#                         else:
#                             input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"P{sp_initial_rw_ex}"))
#                         input_tab.api.AutoFilterMode=False    
#                 except:
#                     pass 

#         print("entry removed successfully")  
#         column_list = input_tab.range("A1").expand('right').value
#         DD_No_column_no = column_list.index('Due Date')+1
#         DD_No_column_letter=num_to_col_letters(Voucher_No_column_no)  
#         Diff_No_column_no = column_list.index('Diff')+1
#         Diff_No_column_letter=num_to_col_letters(Voucher_No_column_no)

#         input_tab.api.Range(f"{Diff_No_column_letter}1").AutoFilter(Field:=f'{Diff_No_column_no}', Criteria1:=['<>0'] ,Operator:=1, Criteria2:=['<>'])

#         input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['<>Total'])

#         dict1={f">{datetime.strptime(input_date,'%m.%d.%Y')}":[DD_No_column_no,DD_No_column_letter,"E","l","K"],f"<={datetime.strptime(input_date,'%m.%d.%Y')-timedelta(days=91)}":[DD_No_column_no,DD_No_column_letter,"E","P","K"]}
#         for key, value in dict1.items():
#             try:
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
#                 time.sleep(1)
#                 sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                 sp_address= input_tab.api.Range(f"{value[2]}2:{value[2]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                 sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#                 input_tab.range(f"{value[3]}{sp_initial_rw}").value = f'=+{value[4]}{sp_initial_rw}'
#                 input_tab.api.Range(f"{value[3]}{sp_initial_rw}:{value[3]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#                 wb.app.api.Selection.FillDown()
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}')
#             except:
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}')
#                 pass  


#         input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}')
#         input_tab.api.AutoFilterMode=False  

#         input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['Total'])

#         sp_lst_row = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"B2:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
#         row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])
#         while row_range[-1]!=sp_lst_row:
#                     sp_lst_row = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
#                     sp_address= input_tab.api.Range(f"B{row_range[-1]}:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#                     sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
#                     row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
#         row_range = sorted(list(set(row_range)))          
#         row_range.insert(0,2)
#         for index,value in enumerate(row_range):
#             if index==0:
#                 inital_value = value
#             else: 
#                 if index>0 and index!=len(row_range)-1:
#                     inital_value = inital_value+1 
#                 if index==len(row_range)-1:
#                     inital_value = row_range[0]     
#                 # if input_tab.range(f"K{value}").value!=None:
#                 input_tab.range(f"K{value}").value = f'=+SUM(K{inital_value}:K{value-1})'

#                 # if input_tab.range(f"L{value}").value!=None:
#                 input_tab.range(f"L{value}").value = f'=+SUM(L{inital_value}:L{value-1})'

#                 # if input_tab.range(f"M{value}").value!=None:
#                 input_tab.range(f"M{value}").value = f'=+SUM(M{inital_value}:M{value-1})'

#                 # if input_tab.range(f"N{value}").value!=None:
#                 input_tab.range(f"N{value}").value = f'=+SUM(N{inital_value}:N{value-1})'

#                 # if input_tab.range(f"O{value}").value!=None:
#                 input_tab.range(f"O{value}").value = f'=+SUM(O{inital_value}:O{value-1})'

#                 # if input_tab.range(f"P{value}").value!=None:
#                 input_tab.range(f"P{value}").value = f'=+SUM(P{inital_value}:P{value-1})'
#                 inital_value = value

#         row_range.pop(-1)                      
#         for index,value in enumerate(row_range):
#             if index==0:
#                 inital_value = value
#             else: 
#                 if input_tab.range(f"K{value}").value>0:
#                     print(f"Accounts payables found:{value}")
#                     inital_value = value
#                 else:
#                     print(f"Accounts receivables found:{value}")
#                     print("starting shifting")
#                     shifting_columns = ["P","O","N","M","L"]
#                     for index2,columns in enumerate(shifting_columns):
#                         # if index>0 and index!=len(row_range)-1:
#                         #     inital_value = inital_value+1     
#                         if columns=="L":
#                             print("reached optimum condition")
#                             break
#                         if columns=="P":
#                             input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
#                             input_tab.api.Range(f"{columns}{inital_value+2}")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
#                             wb.app.api.CutCopyMode=False
#                         if input_tab.range(f"{columns}{value}").value>0:
#                             input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
#                             input_tab.api.Range(f"{shifting_columns[index2+1]}{inital_value+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAll,Operation=win32c.Constants.xlNone,SkipBlanks=True)
#                             input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").ClearContents()

#                     inital_value = value

#         input_tab.autofit()
#         input_tab.api.AutoFilterMode=False  

#         wb.app.api.ActiveWindow.SplitRow=1
#         wb.app.api.ActiveWindow.FreezePanes = True

#         lstr_rw = input_tab.range(f'K'+ str(input_tab.cells.last_cell.row)).end('up').row
#         input_tab.range(f"A1:Q{lstr_rw}").unmerge()

#         bulk_tab= temp_wb.sheets["Bulk"]
#         bulk_tab.api.Copy(After=wb.api.Sheets(2))
#         bulk_tab_it = wb.sheets[2]
#         bulk_tab_it.name = "Bulk_Data(IT)"

#         intial_date = bulk_tab_it.range("B3").value.split("To")[0].strip()
#         last_date = bulk_tab_it.range("B3").value.split("To")[1].strip()

#         intial_date_xl = f"01-01-{year}"

#         last_date = f"{month}-{day}-{year}"
#         xl_input_Date = intial_date_xl + f" To " + last_date
#         bulk_tab_it.range("B3").value = xl_input_Date

#         bulk_tab_it.activate()
#         delete_row_end = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it.api.Range(f"B9:N{delete_row_end}").Delete(win32c.DeleteShiftDirection.xlShiftUp)


#         input_tab.activate()
#         input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['='])
#         sp_lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"A2:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 
#         input_tab.api.Range(f"A{sp_initial_rw}:A{sp_lst_row}").Copy(bulk_tab_it.range(f"B100").api)


#         bulk_tab_it.activate()
#         bulk_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
#         bulk_tab_it.range(f"B9").api.EntireRow.Select()
#         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

#         ini = bulk_tab_it.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it.range(f"B{ini}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
#         ini = bulk_tab_it.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row

#         bulk_tab_it.api.Range(f"C8:N{ini}").Select()
#         wb.app.api.Selection.FillDown()

#         bulk_tab_it.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         bulk_tab_it.api.Range(f"B8:B{ini-1}").Font.Size = 9
#         input_tab.activate()
#         input_tab.api.AutoFilterMode=False
#         input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['Total'])
#         sp_lst_row = input_tab.range(f'K'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"K2:K{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 

#         input_tab.api.Range(f"K{sp_initial_rw}:K{sp_lst_row-1}").Copy(bulk_tab_it.range(f"C8").api)
#         input_tab.activate()
        
#         input_tab.api.Range(f"L{sp_initial_rw}:P{sp_lst_row-1}").Copy(bulk_tab_it.range(f"E8").api)

#         bulk_tab_it.range(f"E8:I{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         bulk_tab_it.range(f"E8:I{ini-1}").api.Font.Size = 9
#         bulk_tab_it.range(f"C8:C{ini-1}").api.Font.Size = 9
#         bulk_tab_it.range(f"C8:C{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

#         retry=0
#         while retry < 10:
#             try:
#                 bulk_wb = xw.Book(input_sheet2,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         bs_tab = bulk_wb.sheets[0]   
#         bs_tab.activate()
#         bs_tab.range(f"A1").select()     
#         bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         cell_value = bs_tab.api.Application.ActiveCell.Address.replace("$","")
#         row_value = re.findall("\d+",cell_value)[0] 
#         bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         cell_value2 = bs_tab.api.Application.ActiveCell.Address.replace("$","")
#         row_value2 = re.findall("\d+",cell_value2)[0]
#         bs_tab.api.Range(f"B{row_value}:C{int(row_value2)-1}").Copy(bs_tab.range(f"I1").api)

#         bs_tab.api.Range(f"J1").AutoFilter(Field:=2, Criteria1:=['=0.00'],Operator:=2,Criteria2:="=0.01")
#         sp_lst_row = bs_tab.range(f'I'+ str(bs_tab.cells.last_cell.row)).end('up').row
#         sp_address= bs_tab.api.Range(f"I2:I{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#         bs_tab.range(f"I{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         bs_tab.api.AutoFilterMode=False 
#         time.sleep(1)
#         bs_total = round(sum(bs_tab.range(f"J2").expand('down').value),2)
#         bs_tab.range(f"I2").expand("table").copy(bulk_tab_it.range(f"L8"))
#         bulk_tab_it.activate()
#         bulk_tab_it.autofit()
#         bs_total_row = bulk_tab_it.range(f'C'+ str(bs_tab.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it.range(f"C{bs_total_row}").value = -bs_total
#         #     Cells.Find(What:="accounts receivable", After:=ActiveCell, LookIn:= _
#         # xlFormulas2, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
#         # xlNext, MatchCase:=False, SearchFormat:=False).Activate
#         companny_name1 = bulk_tab_it.range(f"B8:B{ini-1}").value
#         refined_name1 = [" ".join(name.split(" ")[:-1]) for name in companny_name1]
#         bulk_tab_it.range(f"P8").options(transpose=True).value = refined_name1

#         companny_name2= bulk_tab_it.range(f"L8").expand('down').value
#         refined_name2 = [name.strip() for name in companny_name2]
#         bulk_tab_it.range(f"L8").options(transpose=True).value = refined_name2

#         bulk_tab_it.range(f"J8").value = "=XLOOKUP(P8,L:L,M:M,0)"
#         bulk_tab_it.range(f"J8:J{ini-1}").api.Select()
#                 # bulk_tab_it.api.Range(f"C8:N{ini}").Select()
#         wb.app.api.Selection.FillDown()
#         bulk_tab_it.range(f"J8").expand('down').number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         bulk_tab_it.range(f"J8").expand('down').font.size = 9
#         bulk_tab_it.api.Range(f"J8:J{ini-1}").Copy()
#         bulk_tab_it.api.Range(f"J8:J{ini-1}")._PasteSpecial(Paste=-4163)
#         wb.app.api.CutCopyMode=False
#         bulk_tab_it.range(f"L8").expand('down').api.Delete()
#         bulk_tab_it.api.Range(f"L:L").EntireColumn.Insert()

#         bulk_tab_it.range(f"P8").expand("down").api.Copy(bulk_tab_it.range(f"L8").api)
#         bulk_tab_it.range(f"M8").expand('down').clear_contents()
#         bulk_tab_it.range(f"J8").expand("down").api.Copy(bulk_tab_it.range(f"M8").api)

#         bulk_tab_it.api.Range(f"P:P").EntireColumn.Delete()
#         bulk_tab_it.autofit()
#         for i in range(8,int(ini)):
#             if input_tab.range(f"C{i}").value==0:
#                 print(f"deleted customer={input_tab.range(f'B{i}').value} and deleted row={i}")
#                 input_tab.range(f"{i}:{i}").api.Delete()
#                 break
#             else:
#                 pass  
#         bulk_tab2= temp_wb.sheets["Bulk(2)"]
#         bulk_tab2.api.Copy(After=wb.api.Sheets(3))
#         bulk_tab_it2 = wb.sheets[3]
#         bulk_tab_it2.name = "Bulk_Data(IT)(2)"

#         bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
#         brow_value = re.findall("\d+",bcell_value)[0]
#         bulk_tab_it2.range(f"B{int(brow_value)+1}").expand('table').api.Delete()
#         bulk_tab_it2.range("B3").value = xl_input_Date

#         bulk_tab_it2.range(f"B9:J{int(brow_value)-1}").api.Delete()

#         delete_row_end = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#         delete_row_end2 = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it2.range(f"{delete_row_end2}:{delete_row_end2}").insert()
#         bulk_tab_it2.range(f"{delete_row_end2+1}:{delete_row_end+1}").api.Delete()


#         bulk_tab_it.api.Range(f"B8:C{ini-1}").Copy(bulk_tab_it2.range(f"B100").api)


#         bulk_tab_it2.activate()
#         bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Copy()
#         bulk_tab_it2.range(f"B9").api.EntireRow.Select()
#         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

#         ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it2.range(f"B{ini2}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
#         ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row

#         bulk_tab_it2.api.Range(f"D8:J{ini2-1}").Select()
#         wb.app.api.Selection.FillDown()

#         bulk_tab_it2.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)

#         bulk_tab_it2.api.Range(f"B{ini2-1}").Font.Bold = True

#         bulk_tab_it.api.Range(f"E8:I{ini-1}").Copy(bulk_tab_it2.range(f"E8").api)

#         bulk_tab_it2.api.Range(f"J1").Copy()
#         bulk_tab_it2.api.Range(f"C8:C{ini2-2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
#         bulk_tab_it2.api.Range(f"E8:I{ini2-2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
#         wb.app.api.CutCopyMode=False

#         bs_total_row2 = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it2.range(f"C{bs_total_row2}").value = -bs_total
#         companny_name = bulk_tab_it2.range(f"B8:B{ini2-2}").value
#         refined_name = [" ".join(name.split(" ")[:-1]) + " " for name in companny_name]
#         bulk_tab_it2.range(f"B8").options(transpose=True).value = refined_name

#         retry=0
#         while retry < 10:
#             try:
#                 grp_wb = xw.Book(grp_sheet,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 
#         bulk_tab_it2.activate()
#         bulk_tab_it2.api.Range(f"L8").Value="=+XLOOKUP(B8,'[Group_mapping.xlsx]Sheet1'!$A:$A,'[Group_mapping.xlsx]Sheet1'!$B:$B,0)"

#         bulk_tab_it2.api.Range(f"L8:L{ini2-2}").Select()
#         wb.app.api.Selection.FillDown()
#         bulk_tab_it2.api.Range(f"L7").Select()
#         bulk_tab_it2.api.Range(f"L6").Value = "Xlookup"
#         bulk_tab_it2.api.AutoFilterMode=False
#         bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:='=0')
        
#         sp_lst_row = bulk_tab_it2.range(f'L'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#         if sp_lst_row != 8:
#             sp_address= bulk_tab_it2.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#             sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#         else:
#             sp_initial_rw = 8

#         bulk_tab_it2.range(f"L{sp_initial_rw}").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
#         try:
#             bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1)
#         except:
#             pass    
#         font_colour,Interior_colour = conditional_formatting(range=f"L:L",working_sheet=bulk_tab_it2,working_workbook=wb)
#         bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#         sp_lst_row = bulk_tab_it2.range(f'L'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#         sp_address= bulk_tab_it2.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]

#         bulk_tab_it2.range(f"L{sp_initial_rw}:L{sp_lst_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(bulk_tab_it2.range(f"B100").api)
#         grp_cm_list = bulk_tab_it2.range(f"B100").expand('down').value
#         bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         grp_cm_list2 = list(set(grp_cm_list))
#         bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#         val_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#         if len(grp_cm_list2)>0:
#             for i in range(len(grp_cm_list2)):
#                 # if i >0:
#                 #     val_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row-2
#                 bulk_tab_it2.api.Range(f"L7").Select()
#                 bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=[grp_cm_list2[i]])
#                 sp_lst_row = bulk_tab_it2.range(f'L'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#                 sp_address= bulk_tab_it2.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#                 sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
#                 if bulk_tab_it2.range(f"C{sp_initial_rw}").value + bulk_tab_it2.range(f"C{sp_lst_row}").value<0:
#                     # in_rw = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#                     bulk_tab_it2.range(f"{sp_initial_rw}:{sp_lst_row}").api.EntireRow.Copy()
#                     # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                     bulk_tab_it2.range(f"{val_row+3}:{val_row+3}").api.EntireRow.Select()
#                     wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                     bulk_tab_it2.range(f"{sp_initial_rw}:{sp_lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                 else:
#                     print("second case")

#             bulk_tab_it2.api.Cells.FormatConditions.Delete()
#             bulk_tab_it2.api.AutoFilterMode=False
#         bulk_tab_it2.api.Range(f"L:L").EntireColumn.Delete()
#         font_colour,Interior_colour = conditional_formatting2(range=f"C8:C{ini2-2}",working_sheet=bulk_tab_it2,working_workbook=wb)
#         bulk_tab_it2.api.Range(f"C7").AutoFilter(Field:=2, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

#         sp_lst_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         sp_address= bulk_tab_it2.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#         if int(sp_initial_rw)==6:
#             pass
#         elif int(sp_lst_row) ==int(sp_initial_rw):
#             bulk_tab_it2.range(f"B{sp_initial_rw}").expand("right").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(bulk_tab_it2.range(f"B100").api)
#         else:    
#             bulk_tab_it2.range(f"B{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(bulk_tab_it2.range(f"B100").api)


#         # value_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row

#         bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Copy()
#         bulk_tab_it2.range(f"A{val_row+3}").api.EntireRow.Select()
#         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#         wb.app.api.CutCopyMode=False

#         rw_faltu=bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         if rw_faltu==6:
#             pass
#         elif val_row+3 ==rw_faltu:
#             rw_faltu=bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#             bulk_tab_it2.range(f"B{rw_faltu}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         else:    
#             bulk_tab_it2.range(f"B{rw_faltu}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)

#         if int(sp_initial_rw)==6:
#             pass
#         elif int(sp_lst_row) ==int(sp_initial_rw):
#             bulk_tab_it2.range(f"B{sp_initial_rw}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         else:    
#             bulk_tab_it2.range(f"B{sp_initial_rw}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         bulk_tab_it2.api.AutoFilterMode=False

#         retry=0
#         while retry < 10:
#             try:
#                 company_wb = xw.Book(input_sheet3,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         company_sheet = company_wb.sheets[0] 
#         company_names = company_sheet.range(f"A2").expand('down').value
#         company_names = [names.strip() for names in company_names]
#         company_sheet.range(f"A2").expand('down').api.Copy(bulk_tab_it2.range(f"B100").api)
#         bulk_tab_it2.api.Cells.FormatConditions.Delete()
#         bulk_tab_it2.activate()
#         font_colour,Interior_colour = conditional_formatting(range=f"B:B",working_sheet=bulk_tab_it2,working_workbook=wb)
#         bulk_tab_it2.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

#         sp_lst_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#         sp_address= bulk_tab_it2.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

#         value_row2 = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row

#         bulk_tab_it2.range(f"B{sp_initial_rw}").expand('table').api.Copy(bulk_tab_it2.range(f"B150").api)

#         bulk_tab_it2.range(f"B150").expand('table').api.EntireRow.Copy()
#         # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#         if bulk_tab_it2.range(f"B{value_row2}").value=='Total':
#             value_row2 = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row+2
#         bulk_tab_it2.range(f"A{value_row2+1}").api.EntireRow.Select()
#         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#         bulk_tab_it2.range(f"B{sp_initial_rw}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         bulk_tab_it2.api.AutoFilterMode=False
#         bulk_tab_it2.api.Cells.FormatConditions.Delete()

#         faltu_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it2.range(f"b{faltu_row}").expand('table').api.Delete()
#         faltu_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         bulk_tab_it2.range(f"b{faltu_row}").expand('table').api.Delete()

#         input_tab.api.AutoFilterMode=False

#         raw_df.fillna(0,inplace= True)
#         raw_df = raw_df[raw_df.Customer.isin(company_names) == False]
#         grp_df = raw_df.groupby(['Customer'], sort=False)['Balance','< 10','11 - 30','31 - 60','61 - 90','> 90'].sum().reset_index()
#         grp_df.insert(2,"> 10",grp_df[['11 - 30','31 - 60','61 - 90','> 90']].sum(axis=1))
#         grp_df['As Per BS'] = grp_df['Balance'] - grp_df['< 10'] - grp_df['> 10']
#         grp_df['Customer'] = grp_df['Customer'] + f" "

#         bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
#         brow_value = re.findall("\d+",bcell_value)[0]

#         bulk_tab_it2.api.Range(f"B{int(brow_value)+1}:B{int(brow_value)+len(grp_df)}").EntireRow.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#         bulk_tab_it2.range(f'B{int(brow_value)+1}').options(index = False,header=False).value = grp_df 

#         bulk_tab_it2.range(f'B{int(brow_value)+1}').expand('down').font.bold= False


#         bulk_tab_it2.range(f"B8:J{int(brow_value)-1}").api.Sort(Key1=bulk_tab_it2.range(f"B8:B{int(brow_value)-1}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
      
#         bulk_tab_it2.range(f'B{int(brow_value)+1}').expand('table').api.Sort(Key1=bulk_tab_it2.range(f'B{int(brow_value)+1}').expand('down').api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
            
#         for i in range(len(grp_df['Customer'])):
#             conditional_formatting(range=bulk_tab_it2.range(f'B8').expand('table').get_address(),working_sheet=bulk_tab_it2,working_workbook=wb)
#             bulk_tab_it2.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#             bulk_tab_it2.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=[grp_df['Customer'][i]])
#             sp_lst_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#             sp_address= bulk_tab_it2.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#             sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
#             int_check = bulk_tab_it2.range(f"B{sp_initial_rw}").expand("table").get_address().split(":")[-1]
#             lst_row = re.findall("\d+",int_check .replace("$","").split(":")[0])[0]
#             if bulk_tab_it2.range(f"C{sp_initial_rw}").value + bulk_tab_it2.range(f"C{lst_row}").value<=1:
#                 bulk_tab_it2.range(f"{lst_row}:{lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                 in_rw = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#                 bulk_tab_it2.range(f"{sp_initial_rw}:{sp_initial_rw}").api.EntireRow.Copy()
#                 # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                 bulk_tab_it2.range(f"{in_rw+1}:{in_rw+1}").api.EntireRow.Select()
#                 wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                 bulk_tab_it2.range(f"{sp_initial_rw}:{sp_initial_rw}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                 bulk_tab_it2.api.AutoFilterMode=False
#                 bulk_tab_it2.api.Cells.FormatConditions.Delete()
#             else:
#                 print("second case")
#                 bulk_tab_it2.api.AutoFilterMode=False
#                 bulk_tab_it2.api.Cells.FormatConditions.Delete()

#         #ineligible accounts check
#         bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
#         brow_value = re.findall("\d+",bcell_value)[0]
       
#         if bulk_tab_it2.range(f"B{int(brow_value)+1}").value!=None:
#             pass
#         else:
#             bulk_tab_it2.range(f"{brow_value}:{brow_value}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)


#         #updating formula

#         formula_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row

#         pre_row = bulk_tab_it2.range(f"C{formula_row}").end('up').row

#         fst_rng = bulk_tab_it2.range(f"C8").expand("down").get_address().replace("$","")

#         mid_range = bulk_tab_it2.range(f"C{formula_row}").formula.split("+")[-1].split("-")[0]

#         bulk_tab_it2.range(f"C{formula_row}").formula = f"=+SUM({fst_rng})+{mid_range}-C{pre_row}"

#         input_tab.activate()
#         input_tab.api.Range(f"A:A").EntireColumn.Insert() 
#         initial_tab.activate()
#         initial_tab.cells.unmerge()
#         input_tab.activate()
#         input_tab.api.Range(f"A2").Formula= f"=+XLOOKUP(C2,{initial_tab.name}!C:C,{initial_tab.name}!A:A,0)"
#         st_rw = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row

#         input_tab.api.Range(f"A2:A{st_rw}").Select()
#         wb.app.api.Selection.FillDown()
#         input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=["=0"])
#         input_tab.range("A2").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
#         input_tab.api.Range(f"A1").AutoFilter(Field:=1)
#         input_tab.api.Range(f"A:A").Copy()
#         input_tab.api.Range(f"A:A")._PasteSpecial(Paste=-4163)
#         wb.app.api.CutCopyMode=False

#         tablist={input_tab:win32c.ThemeColor.xlThemeColorAccent2,bulk_tab_it:win32c.ThemeColor.xlThemeColorAccent6,bulk_tab_it2:win32c.ThemeColor.xlThemeColorLight2}
#         for tab,color in tablist.items():
#                 tab.activate()
#                 tab.api.Tab.ThemeColor = color
#                 tab.autofit()
#                 tab.range(f"A1").select()
#         initial_tab.activate()
#         initial_tab.range(f"A1").select()
#         wb.save(f"{output_location}\\AR Aging Bulk {month}{day}-updated"+'.xlsx') 
#         try:
#             wb.app.quit()
#         except:
#             wb.app.quit()  
#         return f"{job_name} Report for {input_date} generated succesfully"

#     except Exception as e:
#         wb.app.kill()
#         raise e
#     finally:
#         try:
#             wb.app.quit()
#         except:
#             pass


# def unbilled_ar(input_date, output_date):
#     try:     
#         job_name = 'Unbilled_AR_automation'
#         month = input_date.split(".")[0]
#         day = input_date.split(".")[1]
#         year = input_date.split(".")[2]
#         dt = datetime.strptime(input_date,"%m.%d.%Y")
#         next_month = (dt.replace(day=1) + timedelta(days=32)).replace(day=1)
#         pre_check = dt.replace(day=1)
#         input_sheet= r'J:\India\BBR\IT_BBR\Reports\Unbilled_AR\Input'+f'\\Unbilled AR {month}{day}.xlsx'
#         output_location = r'J:\India\BBR\IT_BBR\Reports\Unbilled_AR\Output' 
#         master_file = r'\\Bio-India-FS\India Sync$\India\Hamilton\Temporary\BBR Working' + f'\\BBR Master.xlsx'
#         if not os.path.exists(input_sheet):
#             return(f"{input_sheet} Excel file not present for date {input_date}")                 
#         retry=0
#         while retry < 10:
#             try:
#                 wb = xw.Book(input_sheet,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         initial_tab= wb.sheets[0]
#         initial_tab.api.Copy(After=wb.api.Sheets(1))
#         input_tab = wb.sheets[1]
        
#         input_tab.name = "Updated_Data(IT)"
#         input_tab.cells.unmerge()
#         check_column= input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         if check_column ==1:
#                 input_tab.api.Range(f"A:A").EntireColumn.Delete()   

#         input_tab.api.Range(f"1:5").EntireRow.Delete()
#         input_tab.api.Range(f"2:2").EntireRow.Delete()
#         input_tab.autofit()
#         # input_tab.api.Range(f"2:2").EntireRow.Delete()
#         input_tab.activate()


#         column_list = input_tab.range("A1").expand('right').value
#         bldate_No_column_no = column_list.index('B/L date')+1
#         bldate_No_column_letter=num_to_col_letters(bldate_No_column_no)
#         due_date_column_no = column_list.index('Due Date')+1
#         due_date_column_letter=num_to_col_letters(due_date_column_no)
#         date_column_no = column_list.index('Date')+1
#         date_column_letter=num_to_col_letters(date_column_no)        
#         last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
#         dict1={"=":[bldate_No_column_no,bldate_No_column_letter,"A"],f">{datetime.strptime(input_date,'%m.%d.%Y')}":[bldate_No_column_no,bldate_No_column_letter,"L"],f"<{datetime.strptime(input_date,'%m.%d.%Y')}":[date_column_no,date_column_letter,"D"]}
#         for key, value in dict1.items():
#             try:
#                 if key==f">{datetime.strptime(input_date,'%m.%d.%Y')}" or f"<{datetime.strptime(input_date,'%m.%d.%Y')}":
#                     input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
#                 else:
#                     input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key], Operator:=7)
#                 time.sleep(1)
#                 sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                 sp_address= input_tab.api.Range(f"{value[2]}2:{value[2]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                 sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#                 if int(sp_lst_row)!=1:
#                     input_tab.api.Range(f"{sp_initial_rw}:{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#                     time.sleep(1)
#                     wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                     time.sleep(1)
#                 input_tab.api.AutoFilterMode=False   
#             except:
#                 input_tab.api.AutoFilterMode=False 
#                 pass  

#         lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         a = input_tab.range(f"A2:A{lst_rw}").value
#         b = [str(no).strip() for no in a]
#         # try:
#         #     b = [int(str(no).strip().split("#")[1].strip().split(" ")[0]) for no in a]
#         # except:
#         #     b = [str(no).strip().split("#")[1].strip().split(" ")[0] if no!=None else input_tab.api.Range(f"C{index+2}").Value for index,no in enumerate(a) ]
#         #     messagebox.showerror("Invoice Number Error", f"Please re-enter correct value for invoice numbers",parent=root)
#         #     print("Please check invoice numbers")    
#         input_tab.range(f"A2").options(transpose=True).value = b
#         #removing products not required
#         product_column_no = column_list.index('Product')+1
#         product_column_letter=num_to_col_letters(product_column_no)
#         filter_list = ["Product","Admin","Demurrage","Freight Railcar","Freight Truck","Sand","Taxes","True Up - Sales","="]
#         try:
#             input_tab.api.Range(f"{product_column_letter}1").AutoFilter(Field:=f'{product_column_no}', Criteria1:=filter_list, Operator:=7)
#             time.sleep(1)
#             sp_lst_row = input_tab.range(f'{product_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
#             sp_address= input_tab.api.Range(f"{product_column_letter}2:{product_column_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#             sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#             if int(sp_lst_row)!=1:
#                 input_tab.api.Range(f"{sp_initial_rw}:{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#                 time.sleep(1)
#                 wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                 time.sleep(1)
#             input_tab.api.AutoFilterMode=False   
#         except:
#             input_tab.api.AutoFilterMode=False 
#             pass  

#         lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row

#         input_tab.range(f"A2:{last_column_letter}{lst_rw}").api.Sort(Key1=input_tab.range(f"A2:A{lst_rw}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        
#         Particulars_column_no = column_list.index('Particulars')+1
#         Particulars_column_letter=num_to_col_letters(Particulars_column_no)
#         input_tab.api.Range(f"{Particulars_column_letter}1").AutoFilter(Field:=f'{Particulars_column_no}', Criteria1:=["=*SRE*"],Operator:=win32c.AutoFilterOperator.xlAnd)
#         sp_lst_row = input_tab.range(f'{product_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"{product_column_letter}2:{product_column_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#         if int(sp_lst_row)!=1:
#             thick_bottom_border(cellrange=f"{sp_lst_row}:{sp_lst_row}",working_sheet=input_tab,working_workbook=wb)
#         input_tab.api.AutoFilterMode=False 

#         lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row

#         input_tab.api.Range(f"A{lst_rw+5}").Value = -1
#         Quantity_column_no = column_list.index('Quantity')+1
#         Quantity_column_letter=num_to_col_letters(Quantity_column_no)

#         Amount_column_no = column_list.index('Amount')+1
#         Amount_column_letter=num_to_col_letters(Amount_column_no)   

#         TaxCr_Total_column_no = column_list.index('TaxCr Total')+1
#         TaxCr_Total_column_letter=num_to_col_letters(TaxCr_Total_column_no)  

#         input_tab.api.Range(f"A{lst_rw+5}").Copy()
#         input_tab.api.Range(f"{Quantity_column_letter}2:{Quantity_column_letter}{lst_rw}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
#         input_tab.api.Range(f"A{lst_rw+5}").Copy()
#         if int(sp_lst_row)!=1:
#             input_tab.api.Range(f"{Amount_column_letter}2:{TaxCr_Total_column_letter}{sp_lst_row}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
  
#         input_tab.api.Range(f"A{lst_rw+5}").ClearContents()
#         Price_Type_column_no = column_list.index('Price Type')+1
#         Price_Type_column_letter=num_to_col_letters(Price_Type_column_no) 

#         column_list = input_tab.range("A1").expand('right').value
#         # Customer_Name_column_no = column_list.index('Customer Name')+1
#         list1=["Tax","unbilled AR"]
#         list2=["=+AK2-AP2","=+AF2+AQ2"]
#         # Customer_Name_column_no+=1
#         i=0
#         last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         for values in list1:
#             last_column_letter=num_to_col_letters(Price_Type_column_no)
#             input_tab.api.Range(f"{last_column_letter}1").EntireColumn.Insert()
#             input_tab.range(f"{last_column_letter}1").value = values
#             input_tab.range(f"{last_column_letter}2").value = list2[i]
#             time.sleep(1)
#             input_tab.range(f"{last_column_letter}2").copy(input_tab.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
#             i+=1
#             Price_Type_column_no+=1



#         wb.sheets.add("Updated_Pivot(IT)",after=input_tab)
#         ###logger.info("Clearing contents for new sheet")
#         wb.sheets["Updated_Pivot(IT)"].clear_contents()
#         ws2=wb.sheets["Updated_Pivot(IT)"]
#         ###logger.info("Declaring Variables for columns and rows")
#         last_column = input_tab.range('A1').end('right').last_cell.column
#         last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
#         ###logger.info("Creating Pivot Table")
#         PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'{input_tab.name}\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
#         PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Updated_Pivot(IT)'!R2C2", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
#         # PivotTable.PivotFields('Customer Id').Orientation = win32c.PivotFieldOrientation.xlRowField
#         # PivotTable.PivotFields('Customer Id').Position = 1
#         # PivotTable.PivotFields('Customer Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
#         PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
#         PivotTable.PivotFields('Customer Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
#         # PivotTable.PivotFields('Address').Orientation = win32c.PivotFieldOrientation.xlRowField
#         # PivotTable.PivotFields('Address').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
#         # PivotTable.PivotFields('City').Orientation = win32c.PivotFieldOrientation.xlRowField
#         # PivotTable.PivotFields('City').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
#         # PivotTable.PivotFields('State').Orientation = win32c.PivotFieldOrientation.xlRowField
#         # PivotTable.PivotFields('State').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
#         # PivotTable.PivotFields('Zip Code').Orientation = win32c.PivotFieldOrientation.xlRowField
#         # PivotTable.PivotFields('Zip Code').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
#         ###logger.info("Adding particular Data Field in Pivot Table")
#         PivotTable.PivotFields('unbilled AR').Orientation = win32c.PivotFieldOrientation.xlDataField
#         PivotTable.PivotFields('Tax').Orientation = win32c.PivotFieldOrientation.xlDataField

#         # PivotTable.PivotFields('1 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
#         # # PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
#         # PivotTable.PivotFields('31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
#         # # PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
#         # PivotTable.PivotFields('61 - 90').Orientation = win32c.PivotFieldOrientation.xlDataField
#         # # PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
#         # PivotTable.PivotFields('90+').Orientation = win32c.PivotFieldOrientation.xlDataField
#         time.sleep(1)


#         PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'{input_tab.name}\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
#         PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Updated_Pivot(IT)'!R2C10", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
     
#         PivotTable.PivotFields('Terminal').Orientation = win32c.PivotFieldOrientation.xlRowField
#         PivotTable.PivotFields('Terminal').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)

#         PivotTable.PivotFields('Quantity').Orientation = win32c.PivotFieldOrientation.xlDataField
#         PivotTable.PivotFields('Amount').Orientation = win32c.PivotFieldOrientation.xlDataField

#         formula_frsh = ws2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row


#         retry=0
#         while retry < 10:
#             try:
#                 master_wb = xw.Book(master_file,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         ws2.activate()
#         ws2.api.Range(f"A3").Formula = "=+XLOOKUP(B3,'[BBR Master.xlsx]Bulk - AR Aging Master'!$A:$A,'[BBR Master.xlsx]Bulk - AR Aging Master'!$C:$C,0)"
#         ws2.range(f"A3").copy(ws2.range(f"A3:A{formula_frsh-1}"))

#         master_wb.close()

#         tablist={input_tab:win32c.ThemeColor.xlThemeColorAccent2,ws2:win32c.ThemeColor.xlThemeColorAccent6}
#         for tab,color in tablist.items():
#                 tab.activate()
#                 tab.api.Tab.ThemeColor = color
#                 tab.autofit()
#                 tab.range(f"A1").select()
#         input_tab.activate()
#         input_tab.api.Range(f"{due_date_column_letter}:{due_date_column_letter}").Select()
#         ws2.api.Columns("A:A").ColumnWidth = 54.86
#         wb.app.api.ActiveWindow.FreezePanes = True
#         input_tab.range(f"A1").select()
#         initial_tab.activate()   
#         initial_tab.range(f"A1").select() 
#         wb.save(f"{output_location}\\Unbilled AR {month}{day} - updated"+'.xlsx') 
#         try:
#             wb.app.quit()
#         except:
#             wb.app.quit()  
#         return f"{job_name} Report for {input_date} generated succesfully"

#     except Exception as e:
#         wb.app.kill()
#         raise e
#     finally:
#         try:
#             wb.app.quit()
#         except:
#             pass


# def purchased_ar(input_date, output_date):
#     try:       
#         job_name = 'purchased_ar_automation'
#         month = input_date.split(".")[0]
#         day = input_date.split(".")[1]
#         year = input_date.split(".")[2]
#         input_sheet= r'J:\India\BBR\IT_BBR\Reports\Purchased AR\Input'+f'\\Renewable AR {month}{day}.xlsx'
#         output_location = r'J:\India\BBR\IT_BBR\Reports\Purchased AR\Output' 
#         if not os.path.exists(input_sheet):
#             return(f"{input_sheet} Excel file not present for date {input_date}")                 
#         retry=0
#         while retry < 10:
#             try:
#                 wb = xw.Book(input_sheet,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         initial_tab= wb.sheets[0]
#         initial_tab.api.Copy(After=wb.api.Sheets(1))
#         input_tab = wb.sheets[1]
        
#         input_tab.name = "Updated_Data(IT)"

#         check_column= input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         if check_column ==1:
#                 input_tab.api.Range(f"A:A").EntireColumn.Delete()   

#         input_tab.api.Range(f"1:5").EntireRow.Delete()
#         input_tab.api.Range(f"F:M").EntireColumn.Delete() 
#         input_tab.autofit()
#         input_tab.api.Range(f"2:2").EntireRow.Delete()
#         input_tab.activate()


#         column_list = input_tab.range("A1").expand('right').value
#         Voucher_No_column_no = column_list.index('Voucher No')+1
#         Voucher_No_column_letter=num_to_col_letters(Voucher_No_column_no)
#         last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
#         dict1={"Total":[Voucher_No_column_no,Voucher_No_column_letter,"B"],"=":[Voucher_No_column_no,Voucher_No_column_letter,"A"]}
#         for key, value in dict1.items():
#             try:
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key], Operator:=7)
#                 time.sleep(1)
#                 sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                 sp_address= input_tab.api.Range(f"{value[2]}2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                 sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#                 if int(sp_lst_row)!=1:
#                     input_tab.api.Range(f"{sp_initial_rw}:{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#                     time.sleep(1)
#                     wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                     time.sleep(1)
#                 input_tab.api.AutoFilterMode=False 
#             except:
#                 input_tab.api.AutoFilterMode=False 
#                 pass  

#         input_tab.api.Range(f"C:C").EntireColumn.api.Delete()

#         input_tab.api.Range(f"C:C").TextToColumns(Destination:=input_tab.api.Range("C1"),DataType:=win32c.TextParsingType.xlDelimited,TextQualifier:=win32c.Constants.xlDoubleQuote,Tab:=True,FieldInfo:=[1,3],TrailingMinusNumbers:=True)
#         input_tab.api.Range(f"D:D").TextToColumns(Destination:=input_tab.api.Range("D1"),DataType:=win32c.TextParsingType.xlDelimited,TextQualifier:=win32c.Constants.xlDoubleQuote,Tab:=True,FieldInfo:=[1,3],TrailingMinusNumbers:=True)


#         input_tab.api.Range(f"G:G").EntireColumn.Insert()
#         input_tab.api.Range(f"G1").Value = "Current"
#         input_tab.range(f"M1").value = "Diff"
#         input_tab.range(f"M2").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         input_tab.range(f"M2").value='=+F2-SUM(G2:L2)'
#         lsr_rw = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         input_tab.api.Range(f"{lsr_rw+1}:{lsr_rw+10}").EntireRow.api.Delete()
#         input_tab.api.Range(f"M2:M{lsr_rw}").Select()
#         wb.app.api.Selection.FillDown()
        
#         input_tab.api.AutoFilterMode=False
#         input_tab.api.Range(f"A1:M{lsr_rw}").AutoFilter(Field:=13, Criteria1:=["<>0"])
#         input_tab.api.Range(f"A1:M{lsr_rw}").AutoFilter(Field:=4, Criteria1:=[f'>={datetime.now().date().replace(day=int(day),month=int(month),year=int(year))}'])

#         sp_lst_row = input_tab.range(f'F'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"F2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]       


#         input_tab.api.Range(f"G{sp_initial_rw}").Value = f'=+F{sp_initial_rw}'
#         input_tab.api.Range(f"G{sp_initial_rw}:G{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         wb.app.api.Selection.FillDown()

#         input_tab.api.AutoFilterMode=False

#         lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         input_tab.api.Range(f"G2:G{lst_row}").Copy()
#         input_tab.api.Range(f"G2")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
#         wb.app.api.CutCopyMode=False

#         input_tab.api.Range(f"A1:M{lsr_rw}").AutoFilter(Field:=13, Criteria1:=["<>0"])


#         sp_lst_row = input_tab.range(f'F'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"F2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 
#         input_tab.api.Range(f"L{sp_initial_rw}").Value = f'=+F{sp_initial_rw}'
#         if int(sp_initial_rw)==int(sp_lst_row):
#             pass
#         else:
#             input_tab.api.Range(f"L{sp_initial_rw}:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#             wb.app.api.Selection.FillDown()

#         input_tab.api.AutoFilterMode=False
#         lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         input_tab.api.Range(f"L2:L{lst_row}").Copy()
#         input_tab.api.Range(f"L2")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
#         wb.app.api.CutCopyMode=False
#         input_tab.api.Range(f"M:M").EntireColumn.api.Delete()

#         input_tab.api.Range(f"N1").Value = f'-1'
#         lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         input_tab.api.Range(f"N1").Copy()
#         input_tab.api.Range(f"F2:L{lst_row}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
#         input_tab.range(f"F2:L{lst_row}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

#         input_tab.api.Range(f"N1").api.Delete() 

#         input_tab.api.Range(f"E:E").EntireColumn.Copy()
#         input_tab.api.Range(f"N1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone)

#         input_tab.api.Range(f"E:E").EntireColumn.api.Delete()

#         input_tab.api.Range(f"B:B").EntireColumn.Insert()
#         lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         a = input_tab.range(f"N2:N{lst_row}").value
#         try:
#             b = [int(str(no).strip().split("#")[1].strip().split(" ")[0]) for no in a]
#         except:
#             b = [str(no).strip().split("#")[1].strip().split(" ")[0] if no!=None else input_tab.api.Range(f"C{index+2}").Value for index,no in enumerate(a) ]
#             messagebox.showerror("Invoice Number Error", f"Please re-enter correct value for invoice numbers",parent=root)
#             print("Please check invoice numbers")    
#         input_tab.range(f"C2").options(transpose=True).value = b
#         input_tab.range(f"B2").value = 2
#         input_tab.api.Range(f"B2:B{lst_row}").Select()
#         wb.app.api.Selection.FillDown()
#         input_tab.api.Range(f"1:1").EntireRow.Insert()
#         column_headers = ["Customer Name","Tier","Invoice No.","Posting Date","Due Date","Invoice Amount","Current","'1-10","'11-30","31-60","61-90",">90"]
#         input_tab.range(f"A1").value = column_headers
#         for index,value in enumerate(column_headers):
#             column_index = index+1
#             column_letter=num_to_col_letters(index+1)
#             input_tab.api.Range(f"{column_letter}1").HorizontalAlignment = win32c.Constants.xlCenter
#             input_tab.api.Range(f"{column_letter}1").VerticalAlignment = win32c.Constants.xlCenter
#             input_tab.api.Range(f"{column_letter}1").WrapText = True
#             input_tab.api.Range(f"{column_letter}1").Font.Bold = True
#             input_tab.api.Range(f"{column_letter}1").RowHeight = 65


#         input_tab.range(f"A3:N{lst_row+1}").api.Sort(Key1=input_tab.range(f"A3:A{lst_row+1}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
#         input_tab.api.Tab.ThemeColor = win32c.ThemeColor.xlThemeColorAccent4
#         freezepanes_for_tab(cellrange="3:3",working_sheet=input_tab,working_workbook=wb)
#         wb.save(f"{output_location}\\Renewable AR {month}{day} - updated"+'.xlsx') 
#         try:
#             wb.app.quit()
#         except:
#             wb.app.quit()  
#         return f"{job_name} Report for {input_date} generated succesfully"

#     except Exception as e:
#         wb.app.kill()
#         raise e
#     finally:
#         try:
#             wb.app.quit()
#         except:
#             pass

# def ar_ageing_rack(input_date, output_date):
#     try:
#         today_date=date.today()     
#         job_name = 'ar_ageing_Rack'
#         month = input_date.split(".")[0]
#         day = input_date.split(".")[1]
#         year = input_date.split(".")[-1]
#         input_sheet= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Input'+f'\\AR Aging Rack {month}{day}.xlsx'
#         output_location = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Output'
#         input_sheet2= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Input'+f'\\BS Rack {month}{day}.xlsx'
#         input_sheet3= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Template_File'+f'\\Biourja_mapping.xlsx'
#         input_sheet4 = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Template_File'+f'\\AR Aging Rack Template.xlsx'
#         grp_sheet = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Template_File'+f'\\Group_mapping.xlsx'
#         if not os.path.exists(input_sheet):
#             return(f"{input_sheet} Excel file not present for date {input_date}")  
#         if not os.path.exists(input_sheet2):
#             return(f"{input_sheet2} Excel file not present for date {input_date}")  
#         if not os.path.exists(input_sheet3):
#             return(f"{input_sheet3} Excel file not present")    
#         if not os.path.exists(input_sheet4):
#             return(f"{input_sheet4} Excel file not present")                       
#         raw_df = pd.read_excel(input_sheet,skiprows=[0,1,2,3,4,5])    
#         # raw_df = raw_df[(raw_df[raw_df.columns[0]] == 'Demurrage')]
#         # raw_df = raw_df.iloc[:,[0,1,-6,-5,-4,-3,-2,-1]]
#         # raw_df.columns = ['dem_check',"Customer","Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]

#         temp_df = raw_df.loc[:,[raw_df.columns[0],raw_df.columns[1],raw_df.columns[2],raw_df.columns[-6],raw_df.columns[-5],raw_df.columns[-4],raw_df.columns[-3],raw_df.columns[-2],raw_df.columns[-1]]]
#         temp_df = temp_df.dropna(axis=0,subset=[temp_df.columns[1]])
#         t_df = temp_df.reset_index(drop=True)
#         t_df.columns=['dem_check','Date','Due Date',"Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]
#         company_name=''
#         t_df['COMPANY']=''
#         for i,x in t_df.iterrows():
#             try:
#                 print(i,x)
#                 datetime.strptime(x['Date'],'%m-%d-%Y')
#                 t_df['COMPANY'][i]=company_name
#             except:
#                 company_name=x['Date']
#                 print(company_name)
#         t_df = t_df.reindex(columns =['dem_check','COMPANY','Date','Due Date',"Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"])
#         t_df = t_df[(t_df[t_df.columns[0]] == 'Demurrage')]
#         t_df = t_df.reset_index(drop=True)
#         t_df['Date'] = [datetime.strptime(t_df['Date'][x],'%m-%d-%Y') for x in range(len(t_df['Date']))]
#         for i,x in t_df.iterrows():
#             days = (datetime.strptime(input_date,'%m.%d.%Y')-t_df['Due Date'][i]).days
#             if days<=10:
#                 t_df['< 10'][i] = t_df['Balance'][i]
#             elif days>10 and days<=30:
#                 t_df['11 - 30'][i] = t_df['Balance'][i]  
#             elif days>30 and days<=60:
#                 t_df['31 - 60'][i] = t_df['Balance'][i]
#             elif days>60 and days<=90:
#                 t_df['61 - 90'][i] = t_df['Balance'][i]  
#             elif days>90:
#                 t_df['> 90'][i] = t_df['Balance'][i]                 
#             else:
#                 print(f"found new case in demurrange due date:{days} for due date {t_df['Due Date'][i]}")                                                                 
#         # t_df = t_df.iloc[:,[0,1,-6,-5,-4,-3,-2,-1]]
#         # t_df.columns = ['dem_check',"Customer","Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]
#         retry=0
#         while retry < 10:
#             try:
#                 temp_wb = xw.Book(input_sheet4,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e                     
#         retry=0
#         while retry < 10:
#             try:
#                 wb = xw.Book(input_sheet,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         initial_tab= wb.sheets[0]
#         initial_tab.api.Copy(After=wb.api.Sheets(1))
#         input_tab = wb.sheets[1]
        
#         input_tab.name = "Updated_Data(IT)"

#         # check_column= input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         # if check_column ==1:
#         input_tab.api.Range(f"A:A").EntireColumn.Delete()   

#         input_tab.api.Range(f"1:5").EntireRow.Delete()
#         # input_tab.api.Range(f"I:L").EntireColumn.Delete() 
#         input_tab.autofit()
#         input_tab.api.Range(f"2:2").EntireRow.Delete()
#         input_tab.activate()
#         input_tab.cells.unmerge()

#         column_list = input_tab.range("A1").expand('right').value
#         Voucher_No_column_no = column_list.index('Voucher')+1
#         Voucher_No_column_letter=num_to_col_letters(Voucher_No_column_no)
#         last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
#         input_tab.api.AutoFilterMode=False
#         dict1={"<>":[Voucher_No_column_no,Voucher_No_column_letter,"B"]}
#         for key, value in dict1.items():
#             try:
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
#                 time.sleep(1)
#                 sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                 sp_address= input_tab.api.Range(f"{value[2]}2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                 sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#             except:
#                 pass  

#         input_tab.range(f"N1").value = "Diff"
#         input_tab.range(f"N{sp_initial_rw}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         input_tab.range(f"N{sp_initial_rw}").value=f'=+H{sp_initial_rw}-SUM(I{sp_initial_rw}:M{sp_initial_rw})'
#         lsr_rw = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
#         input_tab.api.Range(f"{lsr_rw+1}:{lsr_rw+10}").EntireRow.Delete()
#         input_tab.api.Range(f"N{sp_initial_rw}:N{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#         wb.app.api.Selection.FillDown()
#         input_tab.autofit()
#         freezepanes_for_tab(cellrange="2:2",working_sheet=input_tab,working_workbook=wb)


#         for i in range(2,int(f'{lsr_rw}')):
#             if input_tab.range(f"E{i}").value=="Opb:OPB-911" or input_tab.range(f"F{i}").value=="Opb:OPB-911":
#                 # print(f"deleted customer={input_tab.range(f'A{i}').value} and deleted row={i}")
#                 # input_tab.range(f"{i}:{i}").delete()
#                 input_tab.range(f"B{i}").value = input_tab.range(f"A{i}").value
#                 input_tab.range(f"M{i}").value = input_tab.range(f"H{i}").value
#                 break
#             else:
#                 pass  

#         # input_tab.range(f"Q{sp_initial_rw}:Q{sp_lst_row}")
        
#         # voucher_filters = input_tab.range(f"E2:E{sp_lst_row}").value
#         # jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Jrn" in filter]
#         # input_tab.api.AutoFilterMode=False
#         # if len(jeneral_entry)>0:
#         #     for value in jeneral_entry:
#         #         for index,filter in value.items():
#         #             try:
#         #                 input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
#         #                 time.sleep(1)
#         #                 sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
#         #                 sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#         #                 sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
#         #                 if messagebox.askyesno("Jrn Entry Found!!!",'Do you want this entry to be removed'):
#         #                     print("remove entry") 
#         #                     company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
#         #                     input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").delete()
#         #                     input_tab.api.AutoFilterMode=False 
#         #                     input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
#         #                     sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         #                     sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#         #                     sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
#         #                     length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
#         #                     if length <=1:
#         #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete() 
#         #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete()
#         #                     else:
#         #                         print("Entries found hence no bucket deletion")
#         #                     input_tab.api.AutoFilterMode=False
#         #                 else:
#         #                     print("continue")
#         #                     input_tab.range(f"D{index}").copy(input_tab.range(f"E{index}"))
#         #                     diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{index}").value,"%m-%d-%Y")).days
#         #                     if diff <11:
#         #                         input_tab.range(f"K{index}").copy(input_tab.range(f"L{index}"))
#         #                     elif diff >=11 and diff <31:
#         #                         input_tab.range(f"K{index}").copy(input_tab.range(f"M{index}"))
#         #                     elif diff >=31 and diff <61:
#         #                         input_tab.range(f"K{index}").copy(input_tab.range(f"N{index}"))
#         #                     elif diff >=61 and diff <91:
#         #                         input_tab.range(f"K{index}").copy(input_tab.range(f"O{index}"))
#         #                     else:
#         #                         input_tab.range(f"K{index}").copy(input_tab.range(f"P{index}"))
#         #                     input_tab.api.AutoFilterMode=False    
#         #             except:
#         #                 pass   

#         # jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Exc" in filter]
#         # input_tab.api.AutoFilterMode=False
#         # if len(jeneral_entry)>0:
#         #     for value in jeneral_entry:
#         #         for index,filter in value.items():
#         #             try:
#         #                 input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
#         #                 time.sleep(1)
#         #                 sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
#         #                 sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#         #                 sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
#         #                 if messagebox.askyesno("Exc Entry Found!!!",'Do you want this entry to be removed'):
#         #                     print("remove entry") 
#         #                     company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
#         #                     input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").delete()
#         #                     input_tab.api.AutoFilterMode=False 
#         #                     input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
#         #                     sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         #                     sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#         #                     sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
#         #                     length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
#         #                     if length <=1:
#         #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete() 
#         #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete()
#         #                     else:
#         #                         print("Entries found hence no bucket deletion")
#         #                     input_tab.api.AutoFilterMode=False
#         #                 else:
#         #                     print("continue")
#         #                     input_tab.range(f"D{sp_initial_rw_ex}").copy(input_tab.range(f"E{sp_initial_rw_ex}"))
#         #                     diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{sp_initial_rw_ex}").value,"%m-%d-%Y")).days
#         #                     if diff <11:
#         #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"L{sp_initial_rw_ex}"))
#         #                     elif diff >=11 and diff <31:
#         #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"M{sp_initial_rw_ex}"))
#         #                     elif diff >=31 and diff <61:
#         #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"N{sp_initial_rw_ex}"))
#         #                     elif diff >=61 and diff <91:
#         #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"O{sp_initial_rw_ex}"))
#         #                     else:
#         #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"P{sp_initial_rw_ex}"))
#         #                     input_tab.api.AutoFilterMode=False    
#         #             except:
#         #                 pass 

#         print("entry removed successfully")  
#         column_list = input_tab.range("A1").expand('right').value
#         DD_No_column_no = column_list.index('Due Date')+1
#         DD_No_column_letter=num_to_col_letters(DD_No_column_no)  
#         Diff_No_column_no = column_list.index('Diff')+1
#         Diff_No_column_letter=num_to_col_letters(Diff_No_column_no)
#         input_tab.api.AutoFilterMode=False
#         input_tab.api.Range(f"{Diff_No_column_letter}1").AutoFilter(Field:=f'{Diff_No_column_no}', Criteria1:=['<>0'] ,Operator:=1, Criteria2:=['<>'])

#         input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['<>Total'])

#         dict1={f">{datetime.strptime(input_date,'%m.%d.%Y')}":[DD_No_column_no,DD_No_column_letter,"B","I","H"],f"<={datetime.strptime(input_date,'%m.%d.%Y')-timedelta(days=91)}":[DD_No_column_no,DD_No_column_letter,"B","M","H"]}
#         for key, value in dict1.items():
#             try:
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
#                 time.sleep(1)
#                 sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
#                 sp_address= input_tab.api.Range(f"{value[2]}2:{value[2]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
#                 sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#                 input_tab.range(f"{value[3]}{sp_initial_rw}").value = f'=+{value[4]}{sp_initial_rw}'
#                 input_tab.api.Range(f"{value[3]}{sp_initial_rw}:{value[3]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
#                 wb.app.api.Selection.FillDown()
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}')
#             except:
#                 input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}')
#                 pass  


 
#         input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}')
#         input_tab.api.AutoFilterMode=False 
#         input_tab.api.Range(f"{DD_No_column_letter}1").AutoFilter(Field:=f'{DD_No_column_no}', Criteria1:=['Total'])

#         sp_lst_row = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"B2:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
#         row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])
#         while row_range[-1]!=sp_lst_row:
#                     sp_lst_row = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
#                     sp_address= input_tab.api.Range(f"B{row_range[-1]}:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#                     sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
#                     row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
#         row_range = sorted(list(set(row_range)))          
#         row_range.insert(0,2)
#         for index,value in enumerate(row_range):
#             if index==0:
#                 inital_value = value
#             else: 
#                 if index>0 and index!=len(row_range)-1:
#                     inital_value = inital_value+1 
#                 if index==len(row_range)-1:
#                     inital_value = row_range[0]     
#                 # if input_tab.range(f"K{value}").value!=None:
#                 input_tab.range(f"H{value}").value = f'=+SUM(H{inital_value}:H{value-1})'

#                 # if input_tab.range(f"L{value}").value!=None:
#                 input_tab.range(f"I{value}").value = f'=+SUM(I{inital_value}:I{value-1})'

#                 # if input_tab.range(f"M{value}").value!=None:
#                 input_tab.range(f"J{value}").value = f'=+SUM(J{inital_value}:J{value-1})'

#                 # if input_tab.range(f"N{value}").value!=None:
#                 input_tab.range(f"K{value}").value = f'=+SUM(K{inital_value}:K{value-1})'

#                 # if input_tab.range(f"O{value}").value!=None:
#                 input_tab.range(f"L{value}").value = f'=+SUM(L{inital_value}:L{value-1})'

#                 # if input_tab.range(f"P{value}").value!=None:
#                 input_tab.range(f"M{value}").value = f'=+SUM(M{inital_value}:M{value-1})'
#                 inital_value = value

#         row_range.pop(-1)                      
#         for index,value in enumerate(row_range):
#             if index==0:
#                 inital_value = value
#             else: 
#                 if input_tab.range(f"H{value}").value>0:
#                     print(f"Accounts payables found:{value}")
#                     inital_value = value
#                 else:
#                     print(f"Accounts receivables found:{value}")
#                     print("starting shifting")
#                     shifting_columns = ["M","L","K","J","I"]
#                     for index2,columns in enumerate(shifting_columns):
#                         # if index>0 and index!=len(row_range)-1:
#                         #     inital_value = inital_value+1     
#                         if columns=="I":
#                             print("reached optimum condition")
#                             break
#                         if columns=="M":
#                             input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
#                             input_tab.api.Range(f"{columns}{inital_value+2}")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
#                             wb.app.api.CutCopyMode=False
#                         if input_tab.range(f"{columns}{value}").value>0:
#                             input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
#                             input_tab.api.Range(f"{shifting_columns[index2+1]}{inital_value+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAll,Operation=win32c.Constants.xlNone,SkipBlanks=True)
#                             input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").ClearContents()

#                     inital_value = value

#         input_tab.autofit()
#         input_tab.api.AutoFilterMode=False  

#         wb.app.api.ActiveWindow.SplitRow=1
#         wb.app.api.ActiveWindow.FreezePanes = True

#         lstr_rw = input_tab.range(f'H'+ str(input_tab.cells.last_cell.row)).end('up').row
#         # input_tab.range(f"A1:Q{lstr_rw}").unmerge()

#         rack_tab= temp_wb.sheets["AR Rack"]
#         rack_tab.api.Copy(After=wb.api.Sheets(2))
#         rack_tab_it = wb.sheets[2]
#         rack_tab_it.name = "Rack_Data(IT)"

#         intial_date = rack_tab_it.range("B3").value.split("To")[0].strip()
#         last_date = rack_tab_it.range("B3").value.split("To")[1].strip()

#         intial_date_xl = f"01-01-{year}"

#         last_date = f"{month}-{day}-{year}"
#         xl_input_Date = intial_date_xl + f" To " + last_date
#         rack_tab_it.range("B3").value = xl_input_Date

#         rack_tab_it.activate()



#         # bulk_tab_it = ""
#         # delete_row_end = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
#         rack_tab_it.api.Range(f"B9:J27").Delete(win32c.DeleteShiftDirection.xlShiftUp)


#         input_tab.activate()
#         input_tab.api.Range(f"{DD_No_column_letter}1").AutoFilter(Field:=f'{DD_No_column_no}', Criteria1:=['='])
#         sp_lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"A2:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 
#         input_tab.api.Range(f"A{sp_initial_rw}:A{sp_lst_row}").Copy(rack_tab_it.range(f"B100").api)


#         rack_tab_it.activate()
#         rack_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
#         rack_tab_it.range(f"B9").api.EntireRow.Select()
#         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

        
#         ini = rack_tab_it.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
#         rack_tab_it.range(f"B{ini}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
#         ini_help = rack_tab_it.range(f'J'+ str(input_tab.cells.last_cell.row)).end('up').row
#         ini = rack_tab_it.range(f'B{ini_help}').end('up').row

#         rack_tab_it.api.Range(f"C8:I{ini}").Select()
#         wb.app.api.Selection.FillDown()

#         rack_tab_it.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         rack_tab_it.api.Range(f"B8:B{ini-1}").Font.Size = 9
#         input_tab.activate()
#         input_tab.api.AutoFilterMode=False
#         input_tab.api.Range(f"{DD_No_column_letter}1").AutoFilter(Field:=f'{DD_No_column_no}', Criteria1:=['Total'])
#         sp_lst_row = input_tab.range(f'H'+ str(input_tab.cells.last_cell.row)).end('up').row
#         sp_address= input_tab.api.Range(f"H2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 

#         input_tab.api.Range(f"H{sp_initial_rw}:H{sp_lst_row-1}").Copy(rack_tab_it.range(f"C8").api)
#         input_tab.activate()
        
#         input_tab.api.Range(f"I{sp_initial_rw}:M{sp_lst_row-1}").Copy(rack_tab_it.range(f"E8").api)

#         rack_tab_it.range(f"E8:I{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         rack_tab_it.range(f"E8:I{ini-1}").api.Font.Size = 9
#         rack_tab_it.range(f"C8:C{ini-1}").api.Font.Size = 9
#         rack_tab_it.range(f"C8:C{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'



#         retry=0
#         while retry < 10:
#             try:
#                 bulk_wb = xw.Book(input_sheet2,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         bs_tab = bulk_wb.sheets[0]   
#         bs_tab.activate()
#         bs_tab.range(f"A1").select()     
#         bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         cell_value = bs_tab.api.Application.ActiveCell.Address.replace("$","")
#         row_value = re.findall("\d+",cell_value)[0] 
#         bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         cell_value2 = bs_tab.api.Application.ActiveCell.Address.replace("$","")
#         row_value2 = re.findall("\d+",cell_value2)[0]
#         bs_tab.api.Range(f"B{row_value}:C{int(row_value2)-1}").Copy(bs_tab.range(f"I1").api)

#         bs_tab.api.Range(f"J1").AutoFilter(Field:=2, Criteria1:=['=0.00'],Operator:=2,Criteria2:="=0.01")
#         sp_lst_row = bs_tab.range(f'I'+ str(bs_tab.cells.last_cell.row)).end('up').row
#         sp_address= bs_tab.api.Range(f"I2:I{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#         if int(sp_initial_rw)==1:
#             pass
#         else:
#             bs_tab.range(f"I{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         bs_tab.api.AutoFilterMode=False 
#         time.sleep(1)
#         bs_total = round(sum(bs_tab.range(f"J2").expand('down').value),2)
#         bs_tab.range(f"I2").expand("table").copy(rack_tab_it.range(f"L8"))
#         rack_tab_it.activate()
#         rack_tab_it.autofit()
#         bs_total_row = rack_tab_it.range(f'C{ini_help-1}').end('down').row
#         rack_tab_it.range(f"C{bs_total_row}").value = bs_total
#         #     Cells.Find(What:="accounts receivable", After:=ActiveCell, LookIn:= _
#         # xlFormulas2, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
#         # xlNext, MatchCase:=False, SearchFormat:=False).Activate
#         companny_name1 = rack_tab_it.range(f"B8:B{ini-1}").value
#         refined_name1 = [" ".join(name.split(" ")[:-1]) for name in companny_name1]
#         rack_tab_it.range(f"P8").options(transpose=True).value = refined_name1

#         companny_name2= rack_tab_it.range(f"L8").expand('down').value
#         refined_name2 = [name.strip() for name in companny_name2]
#         rack_tab_it.range(f"L8").options(transpose=True).value = refined_name2

#         rack_tab_it.range(f"J8").value = "=XLOOKUP(P8,L:L,M:M,0)"
#         rack_tab_it.range(f"J8:J{ini-1}").api.Select()
#                 # bulk_tab_it.api.Range(f"C8:N{ini}").Select()
#         wb.app.api.Selection.FillDown()
#         rack_tab_it.range(f"J8").expand('down').number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
#         rack_tab_it.range(f"J8").expand('down').font.size = 9
#         rack_tab_it.api.Range(f"J8:J{ini-1}").Copy()
#         rack_tab_it.api.Range(f"J8:J{ini-1}")._PasteSpecial(Paste=-4163)
#         wb.app.api.CutCopyMode=False
#         rack_tab_it.range(f"L8").expand('table').delete()
#         rack_tab_it.api.Range(f"N:N").EntireColumn.Delete()

#         # bulk_tab_it.range(f"P8").expand("down").api.Copy(bulk_tab_it.range(f"L8").api)
#         # bulk_tab_it.range(f"M8").expand('down').clear_contents()
#         # bulk_tab_it.range(f"J8").expand("down").api.Copy(bulk_tab_it.range(f"M8").api)

#         # bulk_tab_it.api.Range(f"P:P").EntireColumn.Delete()
#         rack_tab_it.autofit()
#         # bulk_tab2= temp_wb.sheets["Bulk(2)"]
#         # bulk_tab2.api.Copy(After=wb.api.Sheets(3))
#         # bulk_tab_it2 = wb.sheets[3]
#         # bulk_tab_it2.name = "Bulk_Data(IT)(2)"

#         # bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         # bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
#         # brow_value = re.findall("\d+",bcell_value)[0]
#         # bulk_tab_it2.range(f"B{int(brow_value)+1}").expand('table').delete()
#         # bulk_tab_it2.range("B3").value = xl_input_Date

#         # bulk_tab_it2.range(f"B9:J{int(brow_value)-1}").delete()

#         # delete_row_end = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#         # delete_row_end2 = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         # bulk_tab_it2.range(f"{delete_row_end2}:{delete_row_end2}").insert()
#         # bulk_tab_it2.range(f"{delete_row_end2+1}:{delete_row_end+1}").delete()


#         # bulk_tab_it.api.Range(f"B8:C{ini-1}").Copy(bulk_tab_it2.range(f"B100").api)


#         # bulk_tab_it2.activate()
#         # bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Copy()
#         # bulk_tab_it2.range(f"B9").api.EntireRow.Select()
#         # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

#         # ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
#         # bulk_tab_it2.range(f"B{ini2}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
#         # ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row

#         # bulk_tab_it2.api.Range(f"D8:J{ini2-1}").Select()
#         # wb.app.api.Selection.FillDown()

#         # bulk_tab_it2.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)

#         # bulk_tab_it2.api.Range(f"B{ini2-1}").Font.Bold = True

#         # bulk_tab_it.api.Range(f"E8:I{ini-1}").Copy(bulk_tab_it2.range(f"E8").api)

#         rack_tab_it.api.Range(f"J1").Copy()
#         rack_tab_it.api.Range(f"C8:C{ini-1}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
#         rack_tab_it.api.Range(f"E8:I{ini-1}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
#         wb.app.api.CutCopyMode=False

#         # bs_total_row2 = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
#         # bulk_tab_it2.range(f"C{bs_total_row2}").value = -bs_total
#         companny_name = rack_tab_it.range(f"B8:B{ini-1}").value
#         refined_name = [" ".join(name.split(" ")[:-1]) + " " for name in companny_name]
#         rack_tab_it.range(f"B8").options(transpose=True).value = refined_name

#         retry=0
#         while retry < 10:
#             try:
#                 grp_wb = xw.Book(grp_sheet,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 
#         rack_tab_it.activate()
#         del_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
#         rack_tab_it.range(f'B{del_row}').expand('table').delete()
#         rack_tab_it.api.Range(f"L8").Value="=+XLOOKUP(B8,'[Group_mapping.xlsx]Sheet1'!$A:$A,'[Group_mapping.xlsx]Sheet1'!$B:$B,0)"

#         rack_tab_it.api.Range(f"L8:L{ini-1}").Select()
#         wb.app.api.Selection.FillDown()
#         rack_tab_it.api.Range(f"L7").Select()
#         rack_tab_it.api.Range(f"L6").Value = "Xlookup"
#         rack_tab_it.api.AutoFilterMode=False
#         rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:='=0')
        
#         sp_lst_row = rack_tab_it.range(f'L'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#         if sp_lst_row != 8:
#             sp_address= rack_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#             sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#         else:
#             sp_initial_rw = 8

#         rack_tab_it.range(f"L{sp_initial_rw}").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
#         try:
#             rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1)
#         except:
#             pass    
#         font_colour,Interior_colour = conditional_formatting(range=f"L:L",working_sheet=rack_tab_it,working_workbook=wb)

#         rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#         sp_lst_row = rack_tab_it.range(f'L'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#         sp_address= rack_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]

#         try:
#             rack_tab_it.range(f"L{sp_initial_rw}:L{sp_lst_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(rack_tab_it.range(f"B100").api)
#         except:
#             pass  

#         if rack_tab_it.range(f"B100").expand('down').value !=None:
#             grp_cm_list = rack_tab_it.range(f"B100").expand('down').value
#         # bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#             grp_cm_list2 = list(set(grp_cm_list))
#             rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#             val_row = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#             if len(grp_cm_list2)>0:
#                 for i in range(len(grp_cm_list2)):
#                     # if i >0:
#                     #     val_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row-2
#                     rack_tab_it.api.Range(f"L7").Select()
#                     rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=[grp_cm_list2[i]])
#                     sp_lst_row = rack_tab_it.range(f'L'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#                     sp_address= rack_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#                     sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
#                     if rack_tab_it.range(f"C{sp_initial_rw}").value + rack_tab_it.range(f"C{sp_lst_row}").value<0:
#                         # in_rw = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
#                         rack_tab_it.range(f"{sp_initial_rw}:{sp_lst_row}").api.EntireRow.Copy()
#                         # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                         rack_tab_it.range(f"{val_row+2}:{val_row+2}").api.EntireRow.Select()
#                         wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                         rack_tab_it.range(f"{sp_initial_rw}:{sp_lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                     else:
#                         print("second case")

#                 rack_tab_it.api.Cells.FormatConditions.Delete()
#                 rack_tab_it.api.AutoFilterMode=False

#         rack_tab_it.api.Range(f"L:L").EntireColumn.Delete()
#         rack_tab_it.activate()
#         font_colour,Interior_colour = conditional_formatting2(range=f"C8:C{ini-1}",working_sheet=rack_tab_it,working_workbook=wb)
#         rack_tab_it.api.Range(f"C7").AutoFilter(Field:=2, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

#         sp_lst_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
#         sp_address= rack_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
#         if int(sp_initial_rw)==6:
#             rack_tab_it.api.Range(f"C7").AutoFilter(Field:=2)
#         elif int(sp_lst_row) ==int(sp_initial_rw):
#             rack_tab_it.range(f"B{sp_initial_rw}").expand("right").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(rack_tab_it.range(f"B100").api)
#         else:    
#             rack_tab_it.range(f"B{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(rack_tab_it.range(f"B100").api)


#         # value_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row
#         if rack_tab_it.range(f"B100").value!=None:
#             rack_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
#             rack_tab_it.range(f"A{val_row+2}").api.EntireRow.Select()
#             wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#             wb.app.api.CutCopyMode=False

#             rw_faltu=rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
#             if val_row+3 ==rw_faltu:
#                 rw_faltu=rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#                 rack_tab_it.range(f"B{rw_faltu}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#             else:    
#                 rack_tab_it.range(f"B{rw_faltu}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)


#         # if int(sp_initial_rw)==6:
#         #     pass
#         # elif int(sp_lst_row) ==int(sp_initial_rw):
#         #     rack_tab_it.range(f"B{sp_initial_rw}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         # else:    
#         #     rack_tab_it.range(f"B{sp_initial_rw}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#         # rack_tab_it.api.AutoFilterMode=False

#         retry=0
#         while retry < 10:
#             try:
#                 company_wb = xw.Book(input_sheet3,update_links=False) 
#                 break
#             except Exception as e:
#                 time.sleep(5)
#                 retry+=1
#                 if retry ==10:
#                     raise e 

#         company_sheet = company_wb.sheets[0] 
#         company_names = company_sheet.range(f"A2").expand('down').value
#         company_names = [names.strip() for names in company_names]
#         company_sheet.range(f"A2").expand('down').api.Copy(rack_tab_it.range(f"B100").api)
#         rack_tab_it.api.Cells.FormatConditions.Delete()
#         rack_tab_it.activate()
#         font_colour,Interior_colour = conditional_formatting(range=f"B:B",working_sheet=rack_tab_it,working_workbook=wb)
#         rack_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

#         sp_lst_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#         sp_address= rack_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#         sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
#         if rack_tab_it.api.Range(f"B{sp_initial_rw}").Value==None:
#             pass
#         else:
#             print("please check for code this condition is new")
#             value_row2 = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').end('up').row

#             rack_tab_it.range(f"B{sp_initial_rw}").expand('table').api.Copy(rack_tab_it.range(f"B150").api)

#             rack_tab_it.range(f"B150").expand('table').api.EntireRow.Copy()
#             # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#             rack_tab_it.range(f"A{value_row2+1}").api.EntireRow.Select()
#             wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#             rack_tab_it.range(f"B{sp_initial_rw}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#             rack_tab_it.api.AutoFilterMode=False
#             rack_tab_it.api.Cells.FormatConditions.Delete()

#         # faltu_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
#         # rack_tab_it.range(f"b{faltu_row}").expand('table').delete()

#         input_tab.api.AutoFilterMode=False
#         rack_tab_it.api.AutoFilterMode=False
#         rack_tab_it.api.Cells.FormatConditions.Delete()
        
#         faltu_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
#         rack_tab_it.range(f"b{faltu_row}").expand('down').delete()

#         t_df.fillna(0,inplace= True)
#         t_df = t_df[t_df.COMPANY.isin(company_names) == False]
#         grp_df = t_df.groupby(['COMPANY'], sort=False)['Balance','< 10','11 - 30','31 - 60','61 - 90','> 90'].sum().reset_index()
#         grp_df.insert(2,"> 10",grp_df[['11 - 30','31 - 60','61 - 90','> 90']].sum(axis=1))
#         grp_df['As Per BS'] = grp_df['Balance'] - grp_df['< 10'] - grp_df['> 10']
#         for i in range(len(grp_df['COMPANY'])):
#             grp_df['COMPANY'][i] = " ".join(grp_df['COMPANY'][i].split(" ")[:-1]) + f" "

#         # bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         # bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
#         check_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#         if rack_tab_it.range(f"B{check_row}").value=='Total':
#             brow_value = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').row + 2
#         else:
#             brow_value = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row + 1 

#         rack_tab_it.api.Range(f"B{int(ini)}:B{int(ini)+len(grp_df)-1}").EntireRow.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#         rack_tab_it.range(f'B{int(ini)}').options(index = False,header=False).value = grp_df 

#         # rack_tab_it.range(f'B{int(brow_value)}').expand('down').font.bold= False


#         rack_tab_it.range(f"B8:J{ini-1}").api.Sort(Key1=rack_tab_it.range(f"B8:B{ini-1}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
      
#         rack_tab_it.range(f'B{int(ini)}').expand('table').api.Sort(Key1=rack_tab_it.range(f'B{int(ini)+1}').expand('down').api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
#         tell_row = rack_tab_it.range(f'B{int(brow_value)}').end('down').row 
#         count = 0 
#         for i in range(len(grp_df['COMPANY'])):
#             conditional_formatting(range=f'B8:B{tell_row}',working_sheet=rack_tab_it,working_workbook=wb)
#             rack_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
#             rack_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=[grp_df['COMPANY'][i]])
#             sp_lst_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#             sp_address= rack_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
#             sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
#             int_check = rack_tab_it.range(f"B{sp_initial_rw}").expand("table").get_address().split(":")[-1]
#             lst_row = re.findall("\d+",int_check .replace("$","").split(":")[0])[0]
#             if rack_tab_it.range(f"C{sp_initial_rw}").value + rack_tab_it.range(f"C{lst_row}").value<=1:
#                 rack_tab_it.range(f"{lst_row}:{lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                 in_rw = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#                 if count>=1:
#                     in_rw = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
#                     rack_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.EntireRow.Copy()
#                     # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                     rack_tab_it.range(f"{in_rw+1}:{in_rw+1}").api.EntireRow.Select()
#                     wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                 else:    
#                     rack_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.EntireRow.Copy()
#                     # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                     rack_tab_it.range(f"{in_rw+2}:{in_rw+2}").api.EntireRow.Select()
#                     wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
#                 rack_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
#                 rack_tab_it.api.AutoFilterMode=False
#                 rack_tab_it.api.Cells.FormatConditions.Delete()
#                 count+=1
#             else:
#                 print("second case")
#                 rack_tab_it.api.AutoFilterMode=False
#                 rack_tab_it.api.Cells.FormatConditions.Delete()

#         #ineligible accounts check
#         # bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
#         # bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
#         # brow_value = re.findall("\d+",bcell_value)[0]
       
#         # if bulk_tab_it2.range(f"B{int(brow_value)+1}").value!=None:
#         #     pass
#         # else:
#         #     bulk_tab_it2.range(f"{brow_value}:{brow_value}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)


#         #updating formula

#         formula_row = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').end('up').row

#         pre_row = rack_tab_it.range(f"C{formula_row}").end('up').row

#         fst_rng = rack_tab_it.range(f"C8").expand("down").get_address().replace("$","")

#         if type(rack_tab_it.range(f"C{formula_row}").end("down").value)==float:

#             rw = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
#             mid_range = rack_tab_it.range(f"C{rw}").expand("down").get_address().replace("$","")
#             rack_tab_it.range(f"C{formula_row}").formula = f"=+C{pre_row}-SUM({fst_rng})-SUM({mid_range})"
#         else:
#             rack_tab_it.range(f"C{formula_row}").formula = f"=+C{pre_row}-SUM({fst_rng})" 

#         input_tab.activate()
#         input_tab.api.Range(f"A:A").EntireColumn.Insert() 
#         initial_tab.activate()
#         initial_tab.cells.unmerge()
#         input_tab.activate()
#         input_tab.api.Range(f"A2").Formula= f"=+XLOOKUP(C2,{initial_tab.name}!C:C,{initial_tab.name}!A:A,0)"
#         st_rw = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row

#         input_tab.api.Range(f"A2:A{st_rw}").Select()
#         wb.app.api.Selection.FillDown()
#         input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=["=0"])
#         input_tab.range("A2").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
#         input_tab.api.Range(f"A1").AutoFilter(Field:=1)
#         input_tab.api.Range(f"A:A").Copy()
#         input_tab.api.Range(f"A:A")._PasteSpecial(Paste=-4163)
#         wb.app.api.CutCopyMode=False

#         tablist={input_tab:win32c.ThemeColor.xlThemeColorAccent2,rack_tab_it:win32c.ThemeColor.xlThemeColorAccent6}
#         for tab,color in tablist.items():
#                 tab.activate()
#                 tab.api.Tab.ThemeColor = color
#                 tab.autofit()
#                 tab.range(f"A1").select()
#         initial_tab.activate()
#         initial_tab.range(f"A1").select()
#         wb.save(f"{output_location}\\AR Aging Rack {month}{day}-updated"+'.xlsx') 
#         try:
#             wb.app.quit()
#         except:
#             wb.app.quit()  
#         return f"{job_name} Report for {input_date} generated succesfully"

#     except Exception as e:
#         wb.app.kill()
#         raise e
#     finally:
#         try:
#             wb.app.quit()
#         except:
#             pass

        
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
    wp_job_ids = {'ABS':1,'Purchased AR Report':purchased_ar,'Ar Ageing Report(Bulk)':ar_ageing_bulk, 'Open Gr':open_gr ,'Ar Ageing Report(Rack)':ar_ageing_rack,'Unbilled AR Report':unbilled_ar,'Cash BBR':bbr_cash,'NLV BBR':bbr_nlv_futures}
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

