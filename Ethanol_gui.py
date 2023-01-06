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

def set_borders(border_range):
    for border_id in range(7,13):
        border_range.api.Borders(border_id).LineStyle=1
        border_range.api.Borders(border_id).Weight=2


def freezepanes_for_tab(cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Rows(cellrange).Select()
        working_workbook.app.api.ActiveWindow.FreezePanes = True
    except Exception as e:
        raise e        

def interior_coloring(colour_value,cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        if working_sheet.api.AutoFilterMode:
            working_sheet.api.Range(cellrange).SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        else:
            working_sheet.api.Range(cellrange).Select()
        a = working_workbook.app.selection.api.Interior
        a.Pattern = win32c.Constants.xlSolid
        a.PatternColorIndex = win32c.Constants.xlAutomatic
        a.Color = colour_value
        a.TintAndShade = 0
        a.PatternTintAndShade = 0        
    except Exception as e:
        raise e  

def conditional_formatting2(range:str,working_sheet,working_workbook):
    try:
        font_colour = -16383844
        Interior_colour = 13551615
        working_sheet.api.Range(range).Select()
        working_workbook.app.selection.api.FormatConditions.Add(Type:=win32c.FormatConditionType.xlCellValue, Operator:=win32c.FormatConditionOperator.xlLess,Formula1:="=0")
        working_workbook.app.selection.api.FormatConditions(working_workbook.app.selection.api.FormatConditions.Count).SetFirstPriority()
        working_workbook.app.selection.api.FormatConditions(1).Font.Color = font_colour
        working_workbook.app.selection.api.FormatConditions(1).Interior.Color = Interior_colour
        working_workbook.app.selection.api.FormatConditions(1).Interior.PatternColorIndex = win32c.Constants.xlAutomatic
        working_workbook.app.selection.api.FormatConditions(1).StopIfTrue = False
        return font_colour,Interior_colour
    except Exception as e:
        raise e

def interior_coloring_by_theme(pattern_tns,tintandshade,colour_value,cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        if working_sheet.api.AutoFilterMode:
            working_sheet.api.Range(cellrange).SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        else:
            working_sheet.api.Range(cellrange).Select()
        a = working_workbook.app.selection.api.Interior
        # a.Pattern = win32c.Constants.xlSolid
        a.PatternColorIndex = win32c.Constants.xlAutomatic
        a.ThemeColor = colour_value
        a.TintAndShade = tintandshade
        a.PatternTintAndShade = pattern_tns    
    except Exception as e:
        raise e  
            

def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e

def insert_all_borders(cellrange:str,working_sheet,working_workbook):
        working_sheet.api.Range(cellrange).Select()
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
        linestylevalues=[win32c.BordersIndex.xlEdgeLeft,win32c.BordersIndex.xlEdgeTop,win32c.BordersIndex.xlEdgeBottom,win32c.BordersIndex.xlEdgeRight,win32c.BordersIndex.xlInsideVertical,win32c.BordersIndex.xlInsideHorizontal]
        for values in linestylevalues:
            a=working_workbook.app.selection.api.Borders(values)
            a.LineStyle = win32c.LineStyle.xlContinuous
            a.ColorIndex = 0
            a.TintAndShade = 0
            a.Weight = win32c.BorderWeight.xlThin

def conditional_formatting(range:str,working_sheet,working_workbook):
    try:
        font_colour = -16383844
        Interior_colour = 13551615
        working_sheet.api.Range(range).Select()
        working_workbook.app.selection.api.FormatConditions.AddUniqueValues()
        working_workbook.app.selection.api.FormatConditions(working_workbook.app.selection.api.FormatConditions.Count).SetFirstPriority()

        working_workbook.app.selection.api.FormatConditions(1).DupeUnique = win32c.DupeUnique.xlDuplicate

        working_workbook.app.selection.api.FormatConditions(1).Font.Color = font_colour
        working_workbook.app.selection.api.FormatConditions(1).Interior.Color = Interior_colour
        working_workbook.app.selection.api.FormatConditions(1).Interior.PatternColorIndex = win32c.Constants.xlAutomatic
        return font_colour,Interior_colour
    except Exception as e:
        raise e

def ar_ageing_bulk(input_date, output_date):
    try:
        today_date=date.today()     
        job_name = 'ar_ageing_automation'
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
        year = input_date.split(".")[-1]
        input_sheet= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing\Input'+f'\\AR Aging Bulk {month}{day}.xlsx'
        output_location = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing\Output'
        input_sheet2= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing\Input'+f'\\BS Bulk {month}{day}.xlsx'
        input_sheet3= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing\Template_File'+f'\\Biourja_mapping.xlsx'
        input_sheet4 = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing\Template_File'+f'\\AR Aging Bulk Template.xlsx'
        grp_sheet = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing\Template_File'+f'\\Group_mapping.xlsx'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")  
        if not os.path.exists(input_sheet2):
            return(f"{input_sheet2} Excel file not present for date {input_date}")  
        if not os.path.exists(input_sheet3):
            return(f"{input_sheet3} Excel file not present")    
        if not os.path.exists(input_sheet4):
            return(f"{input_sheet4} Excel file not present")                       
        raw_df = pd.read_excel(input_sheet)    
        raw_df = raw_df[(raw_df[raw_df.columns[0]] == 'Demurrage')]
        raw_df = raw_df.iloc[:,[0,1,-6,-5,-4,-3,-2,-1]]
        raw_df.columns = ['dem_check',"Customer","Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]
        retry=0
        while retry < 10:
            try:
                temp_wb = xw.Book(input_sheet4,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e                     
        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        initial_tab= wb.sheets[0]
        initial_tab.api.Copy(After=wb.api.Sheets(1))
        input_tab = wb.sheets[1]
        
        input_tab.name = "Updated_Data(IT)"

        # check_column= input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        # if check_column ==1:
        input_tab.api.Range(f"A:A").EntireColumn.Delete()   

        input_tab.api.Range(f"1:5").EntireRow.Delete()
        input_tab.api.Range(f"I:L").EntireColumn.Delete() 
        input_tab.autofit()
        input_tab.api.Range(f"2:2").EntireRow.Delete()
        input_tab.activate()


        column_list = input_tab.range("A1").expand('right').value
        Voucher_No_column_no = column_list.index('Voucher No')+1
        Voucher_No_column_letter=num_to_col_letters(Voucher_No_column_no)
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)

        dict1={"<>":[Voucher_No_column_no,Voucher_No_column_letter,"B"]}
        for key, value in dict1.items():
            try:
                input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
                time.sleep(1)
                sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
                sp_address= input_tab.api.Range(f"{value[2]}2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            except:
                pass  

        input_tab.range(f"Q1").value = "Diff"
        input_tab.range(f"Q{sp_initial_rw}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        input_tab.range(f"Q{sp_initial_rw}").value=f'=+K{sp_initial_rw}-SUM(L{sp_initial_rw}:P{sp_initial_rw})'
        lsr_rw = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"{lsr_rw+1}:{lsr_rw+10}").EntireRow.Delete()
        input_tab.api.Range(f"Q{sp_initial_rw}:Q{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()
        input_tab.autofit()
        freezepanes_for_tab(cellrange="2:2",working_sheet=input_tab,working_workbook=wb)


        for i in range(2,int(f'{lsr_rw}')):
            if input_tab.range(f"B{i}").value=="Opb:OPB-1624" and int(input_tab.range(f"K{i}").value)==58343:
                print(f"deleted customer={input_tab.range(f'A{i}').value} and deleted row={i}")
                input_tab.range(f"{i}:{i}").delete()
                break
            else:
                pass  

        input_tab.range(f"Q{sp_initial_rw}:Q{sp_lst_row}")
        
        voucher_filters = input_tab.range(f"B2:B{sp_lst_row}").value
        jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Jrn" in filter]
        input_tab.api.AutoFilterMode=False
        for value in jeneral_entry:
            for index,filter in value.items():
                try:
                    input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
                    time.sleep(1)
                    sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
                    sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                    sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
                    if messagebox.askyesno("Jrn Entry Found!!!",'Do you want this entry to be removed'):
                        print("remove entry") 
                        company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
                        input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").delete()
                        input_tab.api.AutoFilterMode=False 
                        input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
                        sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
                        sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                        sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
                        length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
                        if length <=1:
                           input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete() 
                           input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete()
                        else:
                            print("Entries found hence no bucket deletion")
                        input_tab.api.AutoFilterMode=False
                    else:
                        print("continue")
                        input_tab.range(f"D{sp_initial_rw_ex}").copy(input_tab.range(f"E{sp_initial_rw_ex}"))
                        diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{sp_initial_rw_ex}").value,"%m-%d-%Y")).days
                        if diff <11:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"L{sp_initial_rw_ex}"))
                        elif diff >=11 and diff <31:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"M{sp_initial_rw_ex}"))
                        elif diff >=31 and diff <61:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"N{sp_initial_rw_ex}"))
                        elif diff >=61 and diff <91:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"O{sp_initial_rw_ex}"))
                        else:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"P{sp_initial_rw_ex}"))
                        input_tab.api.AutoFilterMode=False    
                except:
                    pass   

        jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Exc" in filter]
        input_tab.api.AutoFilterMode=False
        for value in jeneral_entry:
            for index,filter in value.items():
                try:
                    input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
                    time.sleep(1)
                    sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
                    sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                    sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
                    if messagebox.askyesno("Exc Entry Found!!!",'Do you want this entry to be removed'):
                        print("remove entry") 
                        company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
                        input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").delete()
                        input_tab.api.AutoFilterMode=False 
                        input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
                        sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
                        sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                        sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
                        length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
                        if length <=1:
                           input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete() 
                           input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete()
                        else:
                            print("Entries found hence no bucket deletion")
                        input_tab.api.AutoFilterMode=False
                    else:
                        print("continue")
                        input_tab.range(f"D{sp_initial_rw_ex}").copy(input_tab.range(f"E{sp_initial_rw_ex}"))
                        diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{sp_initial_rw_ex}").value,"%m-%d-%Y")).days
                        if diff <11:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"L{sp_initial_rw_ex}"))
                        elif diff >=11 and diff <31:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"M{sp_initial_rw_ex}"))
                        elif diff >=31 and diff <61:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"N{sp_initial_rw_ex}"))
                        elif diff >=61 and diff <91:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"O{sp_initial_rw_ex}"))
                        else:
                            input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"P{sp_initial_rw_ex}"))
                        input_tab.api.AutoFilterMode=False    
                except:
                    pass 

        print("entry removed successfully")  
        column_list = input_tab.range("A1").expand('right').value
        DD_No_column_no = column_list.index('Due Date')+1
        DD_No_column_letter=num_to_col_letters(Voucher_No_column_no)  
        Diff_No_column_no = column_list.index('Diff')+1
        Diff_No_column_letter=num_to_col_letters(Voucher_No_column_no)

        input_tab.api.Range(f"{Diff_No_column_letter}1").AutoFilter(Field:=f'{Diff_No_column_no}', Criteria1:=['<>0'] ,Operator:=1, Criteria2:=['<>'])

        input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['<>Total'])

        dict1={f">{datetime.strptime(input_date,'%m.%d.%Y')}":[DD_No_column_no,DD_No_column_letter,"E","l","K"],f"<={datetime.strptime(input_date,'%m.%d.%Y')-timedelta(days=91)}":[DD_No_column_no,DD_No_column_letter,"E","P","K"]}
        for key, value in dict1.items():
            try:
                input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
                time.sleep(1)
                sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
                sp_address= input_tab.api.Range(f"{value[2]}2:{value[2]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
                input_tab.range(f"{value[3]}{sp_initial_rw}").value = f'=+{value[4]}{sp_initial_rw}'
                input_tab.api.Range(f"{value[3]}{sp_initial_rw}:{value[3]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                wb.app.api.Selection.FillDown()
                input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}')
            except:
                input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}')
                pass  


        input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}')
        input_tab.api.AutoFilterMode=False  

        input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['Total'])

        sp_lst_row = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"B2:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
        row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])
        while row_range[-1]!=sp_lst_row:
                    sp_lst_row = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
                    sp_address= input_tab.api.Range(f"B{row_range[-1]}:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
                    sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
                    row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
        row_range = sorted(list(set(row_range)))          
        row_range.insert(0,2)
        for index,value in enumerate(row_range):
            if index==0:
                inital_value = value
            else: 
                if index>0 and index!=len(row_range)-1:
                    inital_value = inital_value+1 
                if index==len(row_range)-1:
                    inital_value = row_range[0]     
                # if input_tab.range(f"K{value}").value!=None:
                input_tab.range(f"K{value}").value = f'=+SUM(K{inital_value}:K{value-1})'

                # if input_tab.range(f"L{value}").value!=None:
                input_tab.range(f"L{value}").value = f'=+SUM(L{inital_value}:L{value-1})'

                # if input_tab.range(f"M{value}").value!=None:
                input_tab.range(f"M{value}").value = f'=+SUM(M{inital_value}:M{value-1})'

                # if input_tab.range(f"N{value}").value!=None:
                input_tab.range(f"N{value}").value = f'=+SUM(N{inital_value}:N{value-1})'

                # if input_tab.range(f"O{value}").value!=None:
                input_tab.range(f"O{value}").value = f'=+SUM(O{inital_value}:O{value-1})'

                # if input_tab.range(f"P{value}").value!=None:
                input_tab.range(f"P{value}").value = f'=+SUM(P{inital_value}:P{value-1})'
                inital_value = value

        row_range.pop(-1)                      
        for index,value in enumerate(row_range):
            if index==0:
                inital_value = value
            else: 
                if input_tab.range(f"K{value}").value>0:
                    print(f"Accounts payables found:{value}")
                    inital_value = value
                else:
                    print(f"Accounts receivables found:{value}")
                    print("starting shifting")
                    shifting_columns = ["P","O","N","M","L"]
                    for index2,columns in enumerate(shifting_columns):
                        # if index>0 and index!=len(row_range)-1:
                        #     inital_value = inital_value+1     
                        if columns=="L":
                            print("reached optimum condition")
                            break
                        if columns=="P":
                            input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
                            input_tab.api.Range(f"{columns}{inital_value+2}")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
                            wb.app.api.CutCopyMode=False
                        if input_tab.range(f"{columns}{value}").value>0:
                            input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
                            input_tab.api.Range(f"{shifting_columns[index2+1]}{inital_value+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAll,Operation=win32c.Constants.xlNone,SkipBlanks=True)
                            input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").ClearContents()

                    inital_value = value

        input_tab.autofit()
        input_tab.api.AutoFilterMode=False  

        wb.app.api.ActiveWindow.SplitRow=1
        wb.app.api.ActiveWindow.FreezePanes = True

        lstr_rw = input_tab.range(f'K'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.range(f"A1:Q{lstr_rw}").unmerge()

        bulk_tab= temp_wb.sheets["Bulk"]
        bulk_tab.api.Copy(After=wb.api.Sheets(2))
        bulk_tab_it = wb.sheets[2]
        bulk_tab_it.name = "Bulk_Data(IT)"

        intial_date = bulk_tab_it.range("B3").value.split("To")[0].strip()
        last_date = bulk_tab_it.range("B3").value.split("To")[1].strip()

        intial_date_xl = f"01-01-{year}"

        last_date = f"{month}-{day}-{year}"
        xl_input_Date = intial_date_xl + f" To " + last_date
        bulk_tab_it.range("B3").value = xl_input_Date

        bulk_tab_it.activate()
        delete_row_end = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it.api.Range(f"B9:N{delete_row_end}").Delete(win32c.DeleteShiftDirection.xlShiftUp)


        input_tab.activate()
        input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['='])
        sp_lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"A2:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 
        input_tab.api.Range(f"A{sp_initial_rw}:A{sp_lst_row}").Copy(bulk_tab_it.range(f"B100").api)


        bulk_tab_it.activate()
        bulk_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
        bulk_tab_it.range(f"B9").api.EntireRow.Select()
        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

        ini = bulk_tab_it.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it.range(f"B{ini}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
        ini = bulk_tab_it.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row

        bulk_tab_it.api.Range(f"C8:N{ini}").Select()
        wb.app.api.Selection.FillDown()

        bulk_tab_it.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bulk_tab_it.api.Range(f"B8:B{ini-1}").Font.Size = 9
        input_tab.activate()
        input_tab.api.AutoFilterMode=False
        input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['Total'])
        sp_lst_row = input_tab.range(f'K'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"K2:K{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 

        input_tab.api.Range(f"K{sp_initial_rw}:K{sp_lst_row-1}").Copy(bulk_tab_it.range(f"C8").api)
        input_tab.activate()
        
        input_tab.api.Range(f"L{sp_initial_rw}:P{sp_lst_row-1}").Copy(bulk_tab_it.range(f"E8").api)

        bulk_tab_it.range(f"E8:I{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        bulk_tab_it.range(f"E8:I{ini-1}").api.Font.Size = 9
        bulk_tab_it.range(f"C8:C{ini-1}").api.Font.Size = 9
        bulk_tab_it.range(f"C8:C{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        retry=0
        while retry < 10:
            try:
                bulk_wb = xw.Book(input_sheet2,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        bs_tab = bulk_wb.sheets[0]   
        bs_tab.activate()
        bs_tab.range(f"A1").select()     
        bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        cell_value = bs_tab.api.Application.ActiveCell.Address.replace("$","")
        row_value = re.findall("\d+",cell_value)[0] 
        bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        cell_value2 = bs_tab.api.Application.ActiveCell.Address.replace("$","")
        row_value2 = re.findall("\d+",cell_value2)[0]
        bs_tab.api.Range(f"B{row_value}:C{int(row_value2)-1}").Copy(bs_tab.range(f"I1").api)

        bs_tab.api.Range(f"J1").AutoFilter(Field:=2, Criteria1:=['=0.00'],Operator:=2,Criteria2:="=0.01")
        sp_lst_row = bs_tab.range(f'I'+ str(bs_tab.cells.last_cell.row)).end('up').row
        sp_address= bs_tab.api.Range(f"I2:I{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        bs_tab.range(f"I{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bs_tab.api.AutoFilterMode=False 
        time.sleep(1)
        bs_total = round(sum(bs_tab.range(f"J2").expand('down').value),2)
        bs_tab.range(f"I2").expand("table").copy(bulk_tab_it.range(f"L8"))
        bulk_tab_it.activate()
        bulk_tab_it.autofit()
        bs_total_row = bulk_tab_it.range(f'C'+ str(bs_tab.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it.range(f"C{bs_total_row}").value = -bs_total
        #     Cells.Find(What:="accounts receivable", After:=ActiveCell, LookIn:= _
        # xlFormulas2, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        # xlNext, MatchCase:=False, SearchFormat:=False).Activate
        companny_name1 = bulk_tab_it.range(f"B8:B{ini-1}").value
        refined_name1 = [" ".join(name.split(" ")[:-1]) for name in companny_name1]
        bulk_tab_it.range(f"P8").options(transpose=True).value = refined_name1

        companny_name2= bulk_tab_it.range(f"L8").expand('down').value
        refined_name2 = [name.strip() for name in companny_name2]
        bulk_tab_it.range(f"L8").options(transpose=True).value = refined_name2

        bulk_tab_it.range(f"J8").value = "=XLOOKUP(P8,L:L,M:M,0)"
        bulk_tab_it.range(f"J8:J{ini-1}").api.Select()
                # bulk_tab_it.api.Range(f"C8:N{ini}").Select()
        wb.app.api.Selection.FillDown()
        bulk_tab_it.range(f"J8").expand('down').number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        bulk_tab_it.range(f"J8").expand('down').font.size = 9
        bulk_tab_it.api.Range(f"J8:J{ini-1}").Copy()
        bulk_tab_it.api.Range(f"J8:J{ini-1}")._PasteSpecial(Paste=-4163)
        wb.app.api.CutCopyMode=False
        bulk_tab_it.api.Range(f"P:P").EntireColumn.Delete()

        bulk_tab2= temp_wb.sheets["Bulk(2)"]
        bulk_tab2.api.Copy(After=wb.api.Sheets(3))
        bulk_tab_it2 = wb.sheets[3]
        bulk_tab_it2.name = "Bulk_Data(IT)(2)"

        bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
        brow_value = re.findall("\d+",bcell_value)[0]
        bulk_tab_it2.range(f"B{int(brow_value)+1}").expand('table').delete()
        bulk_tab_it2.range("B3").value = xl_input_Date

        bulk_tab_it2.range(f"B9:J{int(brow_value)-1}").delete()

        delete_row_end = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
        delete_row_end2 = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it2.range(f"{delete_row_end2}:{delete_row_end2}").insert()
        bulk_tab_it2.range(f"{delete_row_end2+1}:{delete_row_end+1}").delete()


        bulk_tab_it.api.Range(f"B8:C{ini-1}").Copy(bulk_tab_it2.range(f"B100").api)


        bulk_tab_it2.activate()
        bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Copy()
        bulk_tab_it2.range(f"B9").api.EntireRow.Select()
        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

        ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it2.range(f"B{ini2}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
        ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row

        bulk_tab_it2.api.Range(f"D8:J{ini2-1}").Select()
        wb.app.api.Selection.FillDown()

        bulk_tab_it2.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)

        bulk_tab_it2.api.Range(f"B{ini2-1}").Font.Bold = True

        bulk_tab_it.api.Range(f"E8:I{ini-1}").Copy(bulk_tab_it2.range(f"E8").api)

        bulk_tab_it2.api.Range(f"J1").Copy()
        bulk_tab_it2.api.Range(f"C8:C{ini2-2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        bulk_tab_it2.api.Range(f"E8:I{ini2-2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        wb.app.api.CutCopyMode=False

        bs_total_row2 = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it2.range(f"C{bs_total_row2}").value = -bs_total
        companny_name = bulk_tab_it2.range(f"B8:B{ini2-2}").value
        refined_name = [" ".join(name.split(" ")[:-1]) + " " for name in companny_name]
        bulk_tab_it2.range(f"B8").options(transpose=True).value = refined_name

        retry=0
        while retry < 10:
            try:
                grp_wb = xw.Book(grp_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        bulk_tab_it2.activate()
        bulk_tab_it2.api.Range(f"L8").Value="=+XLOOKUP(B8,'[Group_mapping.xlsx]Sheet1'!$A:$A,'[Group_mapping.xlsx]Sheet1'!$B:$B,0)"

        bulk_tab_it2.api.Range(f"L8:L{ini2-2}").Select()
        wb.app.api.Selection.FillDown()
        bulk_tab_it2.api.Range(f"L7").Select()
        bulk_tab_it2.api.Range(f"L6").Value = "Xlookup"
        bulk_tab_it2.api.AutoFilterMode=False
        bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:='=0')
        
        sp_lst_row = bulk_tab_it2.range(f'L'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
        if sp_lst_row != 8:
            sp_address= bulk_tab_it2.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        else:
            sp_initial_rw = 8

        bulk_tab_it2.range(f"L{sp_initial_rw}").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
        try:
            bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1)
        except:
            pass    
        font_colour,Interior_colour = conditional_formatting(range=f"L:L",working_sheet=bulk_tab_it2,working_workbook=wb)
        bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        sp_lst_row = bulk_tab_it2.range(f'L'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
        sp_address= bulk_tab_it2.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]

        bulk_tab_it2.range(f"L{sp_initial_rw}:L{sp_lst_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(bulk_tab_it2.range(f"B100").api)
        grp_cm_list = bulk_tab_it2.range(f"B100").expand('down').value
        bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        grp_cm_list2 = list(set(grp_cm_list))
        bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        val_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
        if len(grp_cm_list2)>0:
            for i in range(len(grp_cm_list2)):
                # if i >0:
                #     val_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row-2
                bulk_tab_it2.api.Range(f"L7").Select()
                bulk_tab_it2.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=[grp_cm_list2[i]])
                sp_lst_row = bulk_tab_it2.range(f'L'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
                sp_address= bulk_tab_it2.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
                if bulk_tab_it2.range(f"C{sp_initial_rw}").value + bulk_tab_it2.range(f"C{sp_lst_row}").value<0:
                    # in_rw = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
                    bulk_tab_it2.range(f"{sp_initial_rw}:{sp_lst_row}").api.EntireRow.Copy()
                    # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                    bulk_tab_it2.range(f"{val_row+3}:{val_row+3}").api.EntireRow.Select()
                    wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                    bulk_tab_it2.range(f"{sp_initial_rw}:{sp_lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                else:
                    print("second case")

            bulk_tab_it2.api.Cells.FormatConditions.Delete()
            bulk_tab_it2.api.AutoFilterMode=False
        bulk_tab_it2.api.Range(f"L:L").EntireColumn.Delete()
        font_colour,Interior_colour = conditional_formatting2(range=f"C8:C{ini2-2}",working_sheet=bulk_tab_it2,working_workbook=wb)
        bulk_tab_it2.api.Range(f"C7").AutoFilter(Field:=2, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

        sp_lst_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
        sp_address= bulk_tab_it2.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        
        bulk_tab_it2.range(f"B{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(bulk_tab_it2.range(f"B100").api)


        # value_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row

        bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Copy()
        bulk_tab_it2.range(f"A{val_row+3}").api.EntireRow.Select()
        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
        wb.app.api.CutCopyMode=False

        rw_faltu=bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it2.range(f"B{rw_faltu}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bulk_tab_it2.range(f"B{sp_initial_rw}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bulk_tab_it2.api.AutoFilterMode=False

        retry=0
        while retry < 10:
            try:
                company_wb = xw.Book(input_sheet3,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        company_sheet = company_wb.sheets[0] 
        company_names = company_sheet.range(f"A2").expand('down').value
        company_names = [names.strip() for names in company_names]
        company_sheet.range(f"A2").expand('down').api.Copy(bulk_tab_it2.range(f"B100").api)
        bulk_tab_it2.api.Cells.FormatConditions.Delete()
        bulk_tab_it2.activate()
        font_colour,Interior_colour = conditional_formatting(range=f"B:B",working_sheet=bulk_tab_it2,working_workbook=wb)
        bulk_tab_it2.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

        sp_lst_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
        sp_address= bulk_tab_it2.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

        value_row2 = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row

        bulk_tab_it2.range(f"B{sp_initial_rw}").expand('table').api.Copy(bulk_tab_it2.range(f"B150").api)

        bulk_tab_it2.range(f"B150").expand('table').api.EntireRow.Copy()
        # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
        bulk_tab_it2.range(f"A{value_row2+1}").api.EntireRow.Select()
        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
        bulk_tab_it2.range(f"B{sp_initial_rw}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bulk_tab_it2.api.AutoFilterMode=False
        bulk_tab_it2.api.Cells.FormatConditions.Delete()

        faltu_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it2.range(f"b{faltu_row}").expand('table').delete()
        faltu_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
        bulk_tab_it2.range(f"b{faltu_row}").expand('table').delete()

        input_tab.api.AutoFilterMode=False

        raw_df.fillna(0,inplace= True)
        raw_df = raw_df[raw_df.Customer.isin(company_names) == False]
        grp_df = raw_df.groupby(['Customer'], sort=False)['Balance','< 10','11 - 30','31 - 60','61 - 90','> 90'].sum().reset_index()
        grp_df.insert(2,"> 10",grp_df[['11 - 30','31 - 60','61 - 90','> 90']].sum(axis=1))
        grp_df['As Per BS'] = grp_df['Balance'] - grp_df['< 10'] - grp_df['> 10']
        grp_df['Customer'] = grp_df['Customer'] + f" "

        bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
        brow_value = re.findall("\d+",bcell_value)[0]

        bulk_tab_it2.api.Range(f"B{int(brow_value)+1}:B{int(brow_value)+len(grp_df)}").EntireRow.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
        bulk_tab_it2.range(f'B{int(brow_value)+1}').options(index = False,header=False).value = grp_df 

        bulk_tab_it2.range(f'B{int(brow_value)+1}').expand('down').font.bold= False


        bulk_tab_it2.range(f"B8:J{int(brow_value)-1}").api.Sort(Key1=bulk_tab_it2.range(f"B8:B{int(brow_value)-1}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
      
        bulk_tab_it2.range(f'B{int(brow_value)+1}').expand('table').api.Sort(Key1=bulk_tab_it2.range(f'B{int(brow_value)+1}').expand('down').api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
            
        for i in range(len(grp_df['Customer'])):
            conditional_formatting(range=bulk_tab_it2.range(f'B8').expand('table').get_address(),working_sheet=bulk_tab_it2,working_workbook=wb)
            bulk_tab_it2.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
            bulk_tab_it2.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=[grp_df['Customer'][i]])
            sp_lst_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
            sp_address= bulk_tab_it2.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
            int_check = bulk_tab_it2.range(f"B{sp_initial_rw}").expand("table").get_address().split(":")[-1]
            lst_row = re.findall("\d+",int_check .replace("$","").split(":")[0])[0]
            if bulk_tab_it2.range(f"C{sp_initial_rw}").value + bulk_tab_it2.range(f"C{lst_row}").value<=1:
                bulk_tab_it2.range(f"{lst_row}:{lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                in_rw = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
                bulk_tab_it2.range(f"{sp_initial_rw}:{sp_initial_rw}").api.EntireRow.Copy()
                # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                bulk_tab_it2.range(f"{in_rw+1}:{in_rw+1}").api.EntireRow.Select()
                wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                bulk_tab_it2.range(f"{sp_initial_rw}:{sp_initial_rw}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                bulk_tab_it2.api.AutoFilterMode=False
                bulk_tab_it2.api.Cells.FormatConditions.Delete()
            else:
                print("second case")
                bulk_tab_it2.api.AutoFilterMode=False
                bulk_tab_it2.api.Cells.FormatConditions.Delete()

        #ineligible accounts check
        bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
        brow_value = re.findall("\d+",bcell_value)[0]
       
        if bulk_tab_it2.range(f"B{int(brow_value)+1}").value!=None:
            pass
        else:
            bulk_tab_it2.range(f"{brow_value}:{brow_value}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)


        #updating formula

        formula_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row

        pre_row = bulk_tab_it2.range(f"C{formula_row}").end('up').row

        fst_rng = bulk_tab_it2.range(f"C8").expand("down").get_address().replace("$","")

        mid_range = bulk_tab_it2.range(f"C{formula_row}").formula.split("+")[-1].split("-")[0]

        bulk_tab_it2.range(f"C{formula_row}").formula = f"=+SUM({fst_rng})+{mid_range}-C{pre_row}"

        input_tab.activate()
        input_tab.api.Range(f"A:A").EntireColumn.Insert() 
        initial_tab.activate()
        initial_tab.cells.unmerge()
        input_tab.activate()
        input_tab.api.Range(f"A2").Formula= f"=+XLOOKUP(C2,{initial_tab.name}!C:C,{initial_tab.name}!A:A,0)"
        st_rw = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row

        input_tab.api.Range(f"A2:A{st_rw}").Select()
        wb.app.api.Selection.FillDown()
        input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=["=0"])
        input_tab.range("A2").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
        input_tab.api.Range(f"A1").AutoFilter(Field:=1)
        input_tab.api.Range(f"A:A").Copy()
        input_tab.api.Range(f"A:A")._PasteSpecial(Paste=-4163)
        wb.app.api.CutCopyMode=False

        tablist={input_tab:win32c.ThemeColor.xlThemeColorAccent2,bulk_tab_it:win32c.ThemeColor.xlThemeColorAccent6,bulk_tab_it2:win32c.ThemeColor.xlThemeColorLight2}
        for tab,color in tablist.items():
                tab.activate()
                tab.api.Tab.ThemeColor = color
                tab.autofit()
                tab.range(f"A1").select()
        initial_tab.activate()
        initial_tab.range(f"A1").select()
        wb.save(f"{output_location}\\AR Aging Bulk {month}{day}-updated"+'.xlsx') 
        try:
            wb.app.quit()
        except:
            wb.app.quit()  
        return f"{job_name} Report for {input_date} generated succesfully"

    except Exception as e:
        wb.app.kill()
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def purchased_ar(input_date, output_date):
    try:       
        job_name = 'purchased_ar_automation'
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
        year = input_date.split(".")[2]
        input_sheet= r'J:\India\BBR\IT_BBR\Reports\Purchased AR\Input'+f'\\Renewable AR {month}{day}.xlsx'
        output_location = r'J:\India\BBR\IT_BBR\Reports\Purchased AR\Output' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")                 
        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        initial_tab= wb.sheets[0]
        initial_tab.api.Copy(After=wb.api.Sheets(1))
        input_tab = wb.sheets[1]
        
        input_tab.name = "Updated_Data(IT)"

        check_column= input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        if check_column ==1:
                input_tab.api.Range(f"A:A").EntireColumn.Delete()   

        input_tab.api.Range(f"1:5").EntireRow.Delete()
        input_tab.api.Range(f"F:M").EntireColumn.Delete() 
        input_tab.autofit()
        input_tab.api.Range(f"2:2").EntireRow.Delete()
        input_tab.activate()


        column_list = input_tab.range("A1").expand('right').value
        Voucher_No_column_no = column_list.index('Voucher No')+1
        Voucher_No_column_letter=num_to_col_letters(Voucher_No_column_no)
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        dict1={"Total":[Voucher_No_column_no,Voucher_No_column_letter,"B"],"=":[Voucher_No_column_no,Voucher_No_column_letter,"A"]}
        for key, value in dict1.items():
            try:
                input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key], Operator:=7)
                time.sleep(1)
                sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
                sp_address= input_tab.api.Range(f"{value[2]}2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
                input_tab.api.Range(f"{sp_initial_rw}:{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                time.sleep(1)
                wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                time.sleep(1)
                wb.app.api.ActiveSheet.ShowAllData()
            except:
                wb.app.api.ActiveSheet.ShowAllData()
                pass  

        input_tab.api.Range(f"C:C").EntireColumn.Delete()

        input_tab.api.Range(f"C:C").TextToColumns(Destination:=input_tab.api.Range("C1"),DataType:=win32c.TextParsingType.xlDelimited,TextQualifier:=win32c.Constants.xlDoubleQuote,Tab:=True,FieldInfo:=[1,3],TrailingMinusNumbers:=True)
        input_tab.api.Range(f"D:D").TextToColumns(Destination:=input_tab.api.Range("D1"),DataType:=win32c.TextParsingType.xlDelimited,TextQualifier:=win32c.Constants.xlDoubleQuote,Tab:=True,FieldInfo:=[1,3],TrailingMinusNumbers:=True)


        input_tab.api.Range(f"G:G").EntireColumn.Insert()
        input_tab.api.Range(f"G1").Value = "Current"
        input_tab.range(f"M1").value = "Diff"
        input_tab.range(f"M2").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        input_tab.range(f"M2").value='=+F2-SUM(G2:L2)'
        lsr_rw = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"{lsr_rw+1}:{lsr_rw+10}").EntireRow.Delete()
        input_tab.api.Range(f"M2:M{lsr_rw}").Select()
        wb.app.api.Selection.FillDown()
        
        input_tab.api.AutoFilterMode=False
        input_tab.api.Range(f"A1:M{lsr_rw}").AutoFilter(Field:=13, Criteria1:=["<>0"])
        input_tab.api.Range(f"A1:M{lsr_rw}").AutoFilter(Field:=4, Criteria1:=[f'>={datetime.now().date().replace(day=int(day),month=int(month),year=int(year))}'])

        sp_lst_row = input_tab.range(f'F'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"F2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]       


        input_tab.api.Range(f"G{sp_initial_rw}").Value = f'=+F{sp_initial_rw}'
        input_tab.api.Range(f"G{sp_initial_rw}:G{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()

        input_tab.api.AutoFilterMode=False

        lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"G2:G{lst_row}").Copy()
        input_tab.api.Range(f"G2")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
        wb.app.api.CutCopyMode=False

        input_tab.api.Range(f"A1:M{lsr_rw}").AutoFilter(Field:=13, Criteria1:=["<>0"])


        sp_lst_row = input_tab.range(f'F'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"F2:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 
        input_tab.api.Range(f"L{sp_initial_rw}").Value = f'=+F{sp_initial_rw}'
        if int(sp_initial_rw)==int(sp_lst_row):
            pass
        else:
            input_tab.api.Range(f"L{sp_initial_rw}:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
            wb.app.api.Selection.FillDown()

        input_tab.api.AutoFilterMode=False
        lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"L2:L{lst_row}").Copy()
        input_tab.api.Range(f"L2")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
        wb.app.api.CutCopyMode=False
        input_tab.api.Range(f"M:M").EntireColumn.Delete()

        input_tab.api.Range(f"N1").Value = f'-1'
        lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"N1").Copy()
        input_tab.api.Range(f"F2:L{lst_row}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        input_tab.range(f"F2:L{lst_row}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        input_tab.api.Range(f"N1").Delete() 

        input_tab.api.Range(f"E:E").EntireColumn.Copy()
        input_tab.api.Range(f"N1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone)

        input_tab.api.Range(f"E:E").EntireColumn.Delete()

        input_tab.api.Range(f"B:B").EntireColumn.Insert()
        lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        a = input_tab.range(f"N2:N{lst_row}").value
        try:
            b = [int(str(no).strip().split("#")[1].strip().split(" ")[0]) for no in a]
        except:
            b = [str(no).strip().split("#")[1].strip().split(" ")[0] if no!=None else input_tab.api.Range(f"C{index+2}").Value for index,no in enumerate(a) ]
            messagebox.showerror("Invoice Number Error", f"Please re-enter correct value for invoice numbers",parent=root)
            print("Please check invoice numbers")    
        input_tab.range(f"C2").options(transpose=True).value = b
        input_tab.range(f"B2").value = 2
        input_tab.api.Range(f"B2:B{lst_row}").Select()
        wb.app.api.Selection.FillDown()
        input_tab.api.Range(f"1:1").EntireRow.Insert()
        column_headers = ["Customer Name","Tier","Invoice No.","Posting Date","Due Date","Invoice Amount","Current","'1-10","'11-30","31-60","61-90",">90"]
        input_tab.range(f"A1").value = column_headers
        for index,value in enumerate(column_headers):
            column_index = index+1
            column_letter=num_to_col_letters(index+1)
            input_tab.api.Range(f"{column_letter}1").HorizontalAlignment = win32c.Constants.xlCenter
            input_tab.api.Range(f"{column_letter}1").VerticalAlignment = win32c.Constants.xlCenter
            input_tab.api.Range(f"{column_letter}1").WrapText = True
            input_tab.api.Range(f"{column_letter}1").Font.Bold = True
            input_tab.api.Range(f"{column_letter}1").RowHeight = 65


        input_tab.range(f"A3:N{lst_row+1}").api.Sort(Key1=input_tab.range(f"A3:A{lst_row+1}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        input_tab.api.Tab.ThemeColor = win32c.ThemeColor.xlThemeColorAccent4
        freezepanes_for_tab(cellrange="3:3",working_sheet=input_tab,working_workbook=wb)
        wb.save(f"{output_location}\\Renewable AR {month}{day} - updated"+'.xlsx') 
        try:
            wb.app.quit()
        except:
            wb.app.quit()  
        return f"{job_name} Report for {input_date} generated succesfully"

    except Exception as e:
        wb.app.kill()
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


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
    wp_job_ids = {'ABS':1,'Purchased AR Report':purchased_ar,'Ar Ageing Report(Bulk)':ar_ageing_bulk}
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

