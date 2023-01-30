from email.mime import message
import tkinter as tk
from tkinter.filedialog import askdirectory, askopenfilename
from tkinter import N, Menubutton, Tk, StringVar, Text
from tkinter import PhotoImage
from tkinter.font import Font
from tkinter.ttk import Label
from tkinter import Button
from tkinter.ttk import Frame, Style
from tkinter.ttk import OptionMenuss
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




def unbilled_ar(input_date, output_date):
    try:     
        job_name = 'Unbilled_AR_automation'
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
        year = input_date.split(".")[2]
        dt = datetime.strptime(input_date,"%m.%d.%Y")
        next_month = (dt.replace(day=1) + timedelta(days=32)).replace(day=1)
        pre_check = dt.replace(day=1)
        input_sheet= r'J:\India\BBR\IT_BBR\Reports\Unbilled_AR\Input'+f'\\Unbilled AR {month}{day}.xlsx'
        output_location = r'J:\India\BBR\IT_BBR\Reports\Unbilled_AR\Output' 
        master_file = r'\\Bio-India-FS\India Sync$\India\Hamilton\Temporary\BBR Working' + f'\\BBR Master.xlsx'
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
        input_tab.cells.unmerge()
        check_column= input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        if check_column ==1:
                input_tab.api.Range(f"A:A").EntireColumn.Delete()   

        input_tab.api.Range(f"1:5").EntireRow.Delete()
        input_tab.api.Range(f"2:2").EntireRow.Delete()
        input_tab.autofit()
        # input_tab.api.Range(f"2:2").EntireRow.Delete()
        input_tab.activate()


        column_list = input_tab.range("A1").expand('right').value
        bldate_No_column_no = column_list.index('B/L date')+1
        bldate_No_column_letter=num_to_col_letters(bldate_No_column_no)
        due_date_column_no = column_list.index('Due Date')+1
        due_date_column_letter=num_to_col_letters(due_date_column_no)
        date_column_no = column_list.index('Date')+1
        date_column_letter=num_to_col_letters(date_column_no)        
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        dict1={"=":[bldate_No_column_no,bldate_No_column_letter,"A"],f">{datetime.strptime(input_date,'%m.%d.%Y')}":[bldate_No_column_no,bldate_No_column_letter,"L"],f"<{datetime.strptime(input_date,'%m.%d.%Y')}":[date_column_no,date_column_letter,"D"]}
        for key, value in dict1.items():
            try:
                if key==f">{datetime.strptime(input_date,'%m.%d.%Y')}" or f"<{datetime.strptime(input_date,'%m.%d.%Y')}":
                    input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key])
                else:
                    input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key], Operator:=7)
                time.sleep(1)
                sp_lst_row = input_tab.range(f'{value[2]}'+ str(input_tab.cells.last_cell.row)).end('up').row
                sp_address= input_tab.api.Range(f"{value[2]}2:{value[2]}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
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

        lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row
        a = input_tab.range(f"A2:A{lst_rw}").value
        b = [str(no).strip() for no in a]
        # try:
        #     b = [int(str(no).strip().split("#")[1].strip().split(" ")[0]) for no in a]
        # except:
        #     b = [str(no).strip().split("#")[1].strip().split(" ")[0] if no!=None else input_tab.api.Range(f"C{index+2}").Value for index,no in enumerate(a) ]
        #     messagebox.showerror("Invoice Number Error", f"Please re-enter correct value for invoice numbers",parent=root)
        #     print("Please check invoice numbers")    
        input_tab.range(f"A2").options(transpose=True).value = b
        #removing products not required
        product_column_no = column_list.index('Product')+1
        product_column_letter=num_to_col_letters(product_column_no)
        filter_list = ["Product","Admin","Demurrage","Freight Railcar","Freight Truck","Sand","Taxes","True Up - Sales","="]
        try:
            input_tab.api.Range(f"{product_column_letter}1").AutoFilter(Field:=f'{product_column_no}', Criteria1:=filter_list, Operator:=7)
            time.sleep(1)
            sp_lst_row = input_tab.range(f'{product_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
            sp_address= input_tab.api.Range(f"{product_column_letter}2:{product_column_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
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

        lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row

        input_tab.range(f"A2:{last_column_letter}{lst_rw}").api.Sort(Key1=input_tab.range(f"A2:A{lst_rw}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        
        Particulars_column_no = column_list.index('Particulars')+1
        Particulars_column_letter=num_to_col_letters(Particulars_column_no)
        input_tab.api.Range(f"{Particulars_column_letter}1").AutoFilter(Field:=f'{Particulars_column_no}', Criteria1:=["=*SRE*"],Operator:=win32c.AutoFilterOperator.xlAnd)
        sp_lst_row = input_tab.range(f'{product_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"{product_column_letter}2:{product_column_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        if int(sp_lst_row)!=1:
            thick_bottom_border(cellrange=f"{sp_lst_row}:{sp_lst_row}",working_sheet=input_tab,working_workbook=wb)
        input_tab.api.AutoFilterMode=False 

        lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row

        input_tab.api.Range(f"A{lst_rw+5}").Value = -1
        Quantity_column_no = column_list.index('Quantity')+1
        Quantity_column_letter=num_to_col_letters(Quantity_column_no)

        Amount_column_no = column_list.index('Amount')+1
        Amount_column_letter=num_to_col_letters(Amount_column_no)   

        TaxCr_Total_column_no = column_list.index('TaxCr Total')+1
        TaxCr_Total_column_letter=num_to_col_letters(TaxCr_Total_column_no)  

        input_tab.api.Range(f"A{lst_rw+5}").Copy()
        input_tab.api.Range(f"{Quantity_column_letter}2:{Quantity_column_letter}{lst_rw}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        input_tab.api.Range(f"A{lst_rw+5}").Copy()
        if int(sp_lst_row)!=1:
            input_tab.api.Range(f"{Amount_column_letter}2:{TaxCr_Total_column_letter}{sp_lst_row}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
  
        input_tab.api.Range(f"A{lst_rw+5}").ClearContents()
        Price_Type_column_no = column_list.index('Price Type')+1
        Price_Type_column_letter=num_to_col_letters(Price_Type_column_no) 

        column_list = input_tab.range("A1").expand('right').value
        # Customer_Name_column_no = column_list.index('Customer Name')+1
        list1=["Tax","unbilled AR"]
        list2=["=+AK2-AP2","=+AF2+AQ2"]
        # Customer_Name_column_no+=1
        i=0
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(Price_Type_column_no)
            input_tab.api.Range(f"{last_column_letter}1").EntireColumn.Insert()
            input_tab.range(f"{last_column_letter}1").value = values
            input_tab.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            input_tab.range(f"{last_column_letter}2").copy(input_tab.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            Price_Type_column_no+=1



        wb.sheets.add("Updated_Pivot(IT)",after=input_tab)
        ###logger.info("Clearing contents for new sheet")
        wb.sheets["Updated_Pivot(IT)"].clear_contents()
        ws2=wb.sheets["Updated_Pivot(IT)"]
        ###logger.info("Declaring Variables for columns and rows")
        last_column = input_tab.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        ###logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'{input_tab.name}\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Updated_Pivot(IT)'!R2C2", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        # PivotTable.PivotFields('Customer Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Customer Id').Position = 1
        # PivotTable.PivotFields('Customer Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('Address').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Address').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('City').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('City').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('State').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('State').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('Zip Code').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Zip Code').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        ###logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('unbilled AR').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Tax').Orientation = win32c.PivotFieldOrientation.xlDataField

        # PivotTable.PivotFields('1 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('61 - 90').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('90+').Orientation = win32c.PivotFieldOrientation.xlDataField
        time.sleep(1)


        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'{input_tab.name}\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Updated_Pivot(IT)'!R2C10", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
     
        PivotTable.PivotFields('Terminal').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Terminal').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)

        PivotTable.PivotFields('Quantity').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Amount').Orientation = win32c.PivotFieldOrientation.xlDataField

        formula_frsh = ws2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row


        retry=0
        while retry < 10:
            try:
                master_wb = xw.Book(master_file,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        ws2.activate()
        ws2.api.Range(f"A3").Formula = "=+XLOOKUP(B3,'[BBR Master.xlsx]Bulk - AR Aging Master'!$A:$A,'[BBR Master.xlsx]Bulk - AR Aging Master'!$C:$C,0)"
        ws2.range(f"A3").copy(ws2.range(f"A3:A{formula_frsh-1}"))

        master_wb.close()

        tablist={input_tab:win32c.ThemeColor.xlThemeColorAccent2,ws2:win32c.ThemeColor.xlThemeColorAccent6}
        for tab,color in tablist.items():
                tab.activate()
                tab.api.Tab.ThemeColor = color
                tab.autofit()
                tab.range(f"A1").select()
        input_tab.activate()
        input_tab.api.Range(f"{due_date_column_letter}:{due_date_column_letter}").Select()
        ws2.api.Columns("A:A").ColumnWidth = 54.86
        wb.app.api.ActiveWindow.FreezePanes = True
        input_tab.range(f"A1").select()
        initial_tab.activate()   
        initial_tab.range(f"A1").select() 
        wb.save(f"{output_location}\\Unbilled AR {month}{day} - updated"+'.xlsx') 
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