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
# from Common.common import num_to_col_letters
from Common.common import set_borders,freezepanes_for_tab,interior_coloring,conditional_formatting2,interior_coloring_by_theme,num_to_col_letters,insert_all_borders,conditional_formatting,knockOffAmtDiff,row_range_calc,thick_bottom_border





def purchased_ar(input_date, output_date):
    try:   
        root = Tk()    
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
                if int(sp_lst_row)!=1:
                    input_tab.api.Range(f"{sp_initial_rw}:{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                    time.sleep(1)
                    wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                    time.sleep(1)
                input_tab.api.AutoFilterMode=False 
            except:
                input_tab.api.AutoFilterMode=False 
                pass  

        input_tab.Range(f"C:C").EntireColumn.api.Delete()

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
        input_tab.api.Range(f"M:M").EntireColumn.api.Delete()

        input_tab.api.Range(f"N1").Value = f'-1'
        lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"N1").Copy()
        input_tab.api.Range(f"F2:L{lst_row}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        input_tab.range(f"F2:L{lst_row}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        input_tab.api.Range(f"N1").api.Delete() 

        input_tab.api.Range(f"E:E").EntireColumn.Copy()
        input_tab.api.Range(f"N1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone)

        input_tab.api.Range(f"E:E").EntireColumn.api.Delete()

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