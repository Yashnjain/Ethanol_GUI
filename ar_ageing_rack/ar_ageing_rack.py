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





def ar_ageing_rack(input_date, output_date):
    try:
        today_date=date.today()     
        job_name = 'ar_ageing_Rack'
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
        year = input_date.split(".")[-1]
        input_sheet= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Input'+f'\\AR Aging Rack {month}{day}.xlsx'
        output_location = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Output'
        input_sheet2= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Input'+f'\\BS Rack {month}{day}.xlsx'
        input_sheet3= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Template_File'+f'\\Biourja_mapping.xlsx'
        input_sheet4 = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Template_File'+f'\\AR Aging Rack Template.xlsx'
        grp_sheet = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Template_File'+f'\\Group_mapping.xlsx'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")  
        if not os.path.exists(input_sheet2):
            return(f"{input_sheet2} Excel file not present for date {input_date}")  
        if not os.path.exists(input_sheet3):
            return(f"{input_sheet3} Excel file not present")    
        if not os.path.exists(input_sheet4):
            return(f"{input_sheet4} Excel file not present")                       
        raw_df = pd.read_excel(input_sheet,skiprows=[0,1,2,3,4,5])    
        # raw_df = raw_df[(raw_df[raw_df.columns[0]] == 'Demurrage')]
        # raw_df = raw_df.iloc[:,[0,1,-6,-5,-4,-3,-2,-1]]
        # raw_df.columns = ['dem_check',"Customer","Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]

        temp_df = raw_df.loc[:,[raw_df.columns[0],raw_df.columns[1],raw_df.columns[2],raw_df.columns[-6],raw_df.columns[-5],raw_df.columns[-4],raw_df.columns[-3],raw_df.columns[-2],raw_df.columns[-1]]]
        temp_df = temp_df.dropna(axis=0,subset=[temp_df.columns[1]])
        t_df = temp_df.reset_index(drop=True)
        t_df.columns=['dem_check','Date','Due Date',"Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]
        company_name=''
        t_df['COMPANY']=''
        for i,x in t_df.iterrows():
            try:
                print(i,x)
                datetime.strptime(x['Date'],'%m-%d-%Y')
                t_df['COMPANY'][i]=company_name
            except:
                company_name=x['Date']
                print(company_name)
        t_df = t_df.reindex(columns =['dem_check','COMPANY','Date','Due Date',"Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"])
        t_df = t_df[(t_df[t_df.columns[0]] == 'Demurrage')]
        t_df = t_df.reset_index(drop=True)
        t_df['Date'] = [datetime.strptime(t_df['Date'][x],'%m-%d-%Y') for x in range(len(t_df['Date']))]
        for i,x in t_df.iterrows():
            days = (datetime.strptime(input_date,'%m.%d.%Y')-t_df['Due Date'][i]).days
            if days<=10:
                t_df['< 10'][i] = t_df['Balance'][i]
            elif days>10 and days<=30:
                t_df['11 - 30'][i] = t_df['Balance'][i]  
            elif days>30 and days<=60:
                t_df['31 - 60'][i] = t_df['Balance'][i]
            elif days>60 and days<=90:
                t_df['61 - 90'][i] = t_df['Balance'][i]  
            elif days>90:
                t_df['> 90'][i] = t_df['Balance'][i]                 
            else:
                print(f"found new case in demurrange due date:{days} for due date {t_df['Due Date'][i]}")                                                                 
        # t_df = t_df.iloc[:,[0,1,-6,-5,-4,-3,-2,-1]]
        # t_df.columns = ['dem_check',"Customer","Balance","< 10","11 - 30","31 - 60","61 - 90","> 90"]
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
        # input_tab.api.Range(f"I:L").EntireColumn.Delete() 
        input_tab.autofit()
        input_tab.api.Range(f"2:2").EntireRow.Delete()
        input_tab.activate()
        input_tab.cells.unmerge()

        column_list = input_tab.range("A1").expand('right').value
        Voucher_No_column_no = column_list.index('Voucher')+1
        Voucher_No_column_letter=num_to_col_letters(Voucher_No_column_no)
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        input_tab.api.AutoFilterMode=False
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

        input_tab.range(f"N1").value = "Diff"
        input_tab.range(f"N{sp_initial_rw}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        input_tab.range(f"N{sp_initial_rw}").value=f'=+H{sp_initial_rw}-SUM(I{sp_initial_rw}:M{sp_initial_rw})'
        lsr_rw = input_tab.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"{lsr_rw+1}:{lsr_rw+10}").EntireRow.Delete()
        input_tab.api.Range(f"N{sp_initial_rw}:N{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()
        input_tab.autofit()
        freezepanes_for_tab(cellrange="2:2",working_sheet=input_tab,working_workbook=wb)


        for i in range(2,int(f'{lsr_rw}')):
            if input_tab.range(f"E{i}").value=="Opb:OPB-911" or input_tab.range(f"F{i}").value=="Opb:OPB-911":
                # print(f"deleted customer={input_tab.range(f'A{i}').value} and deleted row={i}")
                # input_tab.range(f"{i}:{i}").delete()
                input_tab.range(f"B{i}").value = input_tab.range(f"A{i}").value
                input_tab.range(f"M{i}").value = input_tab.range(f"H{i}").value
                break
            else:
                pass  

        # input_tab.range(f"Q{sp_initial_rw}:Q{sp_lst_row}")
        
        # voucher_filters = input_tab.range(f"E2:E{sp_lst_row}").value
        # jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Jrn" in filter]
        # input_tab.api.AutoFilterMode=False
        # if len(jeneral_entry)>0:
        #     for value in jeneral_entry:
        #         for index,filter in value.items():
        #             try:
        #                 input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
        #                 time.sleep(1)
        #                 sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
        #                 sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        #                 sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
        #                 if messagebox.askyesno("Jrn Entry Found!!!",'Do you want this entry to be removed'):
        #                     print("remove entry") 
        #                     company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
        #                     input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").delete()
        #                     input_tab.api.AutoFilterMode=False 
        #                     input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
        #                     sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        #                     sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        #                     sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
        #                     length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
        #                     if length <=1:
        #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete() 
        #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete()
        #                     else:
        #                         print("Entries found hence no bucket deletion")
        #                     input_tab.api.AutoFilterMode=False
        #                 else:
        #                     print("continue")
        #                     input_tab.range(f"D{index}").copy(input_tab.range(f"E{index}"))
        #                     diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{index}").value,"%m-%d-%Y")).days
        #                     if diff <11:
        #                         input_tab.range(f"K{index}").copy(input_tab.range(f"L{index}"))
        #                     elif diff >=11 and diff <31:
        #                         input_tab.range(f"K{index}").copy(input_tab.range(f"M{index}"))
        #                     elif diff >=31 and diff <61:
        #                         input_tab.range(f"K{index}").copy(input_tab.range(f"N{index}"))
        #                     elif diff >=61 and diff <91:
        #                         input_tab.range(f"K{index}").copy(input_tab.range(f"O{index}"))
        #                     else:
        #                         input_tab.range(f"K{index}").copy(input_tab.range(f"P{index}"))
        #                     input_tab.api.AutoFilterMode=False    
        #             except:
        #                 pass   

        # jeneral_entry =[{index+2:filter} for index,filter in enumerate(voucher_filters) if filter!=None and "Exc" in filter]
        # input_tab.api.AutoFilterMode=False
        # if len(jeneral_entry)>0:
        #     for value in jeneral_entry:
        #         for index,filter in value.items():
        #             try:
        #                 input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=[filter])
        #                 time.sleep(1)
        #                 sp_lst_row_ex = input_tab.range(f'{Voucher_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
        #                 sp_address_Ex= input_tab.api.Range(f"{Voucher_No_column_letter}2:L{sp_lst_row_ex}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        #                 sp_initial_rw_ex = re.findall("\d+",sp_address_Ex.replace("$","").split(":")[0])[0]
        #                 if messagebox.askyesno("Exc Entry Found!!!",'Do you want this entry to be removed'):
        #                     print("remove entry") 
        #                     company_key = input_tab.range(f"A{sp_initial_rw_ex}").value  
        #                     input_tab.range(f"{sp_initial_rw_ex}:{sp_initial_rw_ex}").delete()
        #                     input_tab.api.AutoFilterMode=False 
        #                     input_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=[company_key+f"*"],Operator:=1)
        #                     sp_lst_row_sc = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        #                     sp_address_sc= input_tab.api.Range(f"A2:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        #                     sp_initial_rw_sc = re.findall("\d+",sp_address_sc.replace("$","").split(":")[0])[0]
        #                     length = len(input_tab.api.Range(f"A{sp_initial_rw_sc}:B{sp_lst_row_sc}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Value)
        #                     if length <=1:
        #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete() 
        #                         input_tab.range(f"{sp_initial_rw_sc}:{sp_initial_rw_sc}").delete()
        #                     else:
        #                         print("Entries found hence no bucket deletion")
        #                     input_tab.api.AutoFilterMode=False
        #                 else:
        #                     print("continue")
        #                     input_tab.range(f"D{sp_initial_rw_ex}").copy(input_tab.range(f"E{sp_initial_rw_ex}"))
        #                     diff = (datetime.strptime(input_date,'%m.%d.%Y') - datetime.strptime(input_tab.range(f"D{sp_initial_rw_ex}").value,"%m-%d-%Y")).days
        #                     if diff <11:
        #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"L{sp_initial_rw_ex}"))
        #                     elif diff >=11 and diff <31:
        #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"M{sp_initial_rw_ex}"))
        #                     elif diff >=31 and diff <61:
        #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"N{sp_initial_rw_ex}"))
        #                     elif diff >=61 and diff <91:
        #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"O{sp_initial_rw_ex}"))
        #                     else:
        #                         input_tab.range(f"K{sp_initial_rw_ex}").copy(input_tab.range(f"P{sp_initial_rw_ex}"))
        #                     input_tab.api.AutoFilterMode=False    
        #             except:
        #                 pass 

        print("entry removed successfully")  
        column_list = input_tab.range("A1").expand('right').value
        DD_No_column_no = column_list.index('Due Date')+1
        DD_No_column_letter=num_to_col_letters(DD_No_column_no)  
        Diff_No_column_no = column_list.index('Diff')+1
        Diff_No_column_letter=num_to_col_letters(Diff_No_column_no)
        input_tab.api.AutoFilterMode=False
        input_tab.api.Range(f"{Diff_No_column_letter}1").AutoFilter(Field:=f'{Diff_No_column_no}', Criteria1:=['<>0'] ,Operator:=1, Criteria2:=['<>'])

        input_tab.api.Range(f"{Voucher_No_column_letter}1").AutoFilter(Field:=f'{Voucher_No_column_no}', Criteria1:=['<>Total'])

        dict1={f">{datetime.strptime(input_date,'%m.%d.%Y')}":[DD_No_column_no,DD_No_column_letter,"B","I","H"],f"<={datetime.strptime(input_date,'%m.%d.%Y')-timedelta(days=91)}":[DD_No_column_no,DD_No_column_letter,"B","M","H"]}
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
        input_tab.api.Range(f"{DD_No_column_letter}1").AutoFilter(Field:=f'{DD_No_column_no}', Criteria1:=['Total'])

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
                input_tab.range(f"H{value}").value = f'=+SUM(H{inital_value}:H{value-1})'

                # if input_tab.range(f"L{value}").value!=None:
                input_tab.range(f"I{value}").value = f'=+SUM(I{inital_value}:I{value-1})'

                # if input_tab.range(f"M{value}").value!=None:
                input_tab.range(f"J{value}").value = f'=+SUM(J{inital_value}:J{value-1})'

                # if input_tab.range(f"N{value}").value!=None:
                input_tab.range(f"K{value}").value = f'=+SUM(K{inital_value}:K{value-1})'

                # if input_tab.range(f"O{value}").value!=None:
                input_tab.range(f"L{value}").value = f'=+SUM(L{inital_value}:L{value-1})'

                # if input_tab.range(f"P{value}").value!=None:
                input_tab.range(f"M{value}").value = f'=+SUM(M{inital_value}:M{value-1})'
                inital_value = value

        row_range.pop(-1)                      
        for index,value in enumerate(row_range):
            if index==0:
                inital_value = value
            else: 
                if input_tab.range(f"H{value}").value>0:
                    print(f"Accounts payables found:{value}")
                    inital_value = value
                else:
                    print(f"Accounts receivables found:{value}")
                    print("starting shifting")
                    shifting_columns = ["M","L","K","J","I"]
                    for index2,columns in enumerate(shifting_columns):
                        # if index>0 and index!=len(row_range)-1:
                        #     inital_value = inital_value+1     
                        if columns=="I":
                            print("reached optimum condition")
                            break
                        if columns=="M":
                            input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
                            input_tab.api.Range(f"{columns}{inital_value+2}")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
                            wb.app.api.CutCopyMode=False
                        if input_tab.range(f"{columns}{value}").value>0:   
                            new_column = shifting_columns[shifting_columns.index(columns)-1]
                            while new_column !="M":
                                if input_tab.range(f"{new_column}{value}").value<0:
                                    input_tab.api.Range(f"{new_column}{inital_value+2}:{new_column}{value-1}").Copy() 
                                    input_tab.api.Range(f"{shifting_columns[index2]}{inital_value+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAll,Operation=win32c.Constants.xlNone,SkipBlanks=True)
                                    input_tab.api.Range(f"{new_column}{inital_value+2}:{new_column}{value-1}").ClearContents()
                                    #move those
                                new_column = shifting_columns[shifting_columns.index(new_column)-1]
                            if input_tab.range(f"{columns}{value}").value>0:
                                input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").Copy() 
                                input_tab.api.Range(f"{shifting_columns[index2+1]}{inital_value+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAll,Operation=win32c.Constants.xlNone,SkipBlanks=True)
                                input_tab.api.Range(f"{columns}{inital_value+2}:{columns}{value-1}").ClearContents()


                    inital_value = value

        input_tab.autofit()
        input_tab.api.AutoFilterMode=False  

        wb.app.api.ActiveWindow.SplitRow=1
        wb.app.api.ActiveWindow.FreezePanes = True

        lstr_rw = input_tab.range(f'H'+ str(input_tab.cells.last_cell.row)).end('up').row
        # input_tab.range(f"A1:Q{lstr_rw}").unmerge()

        rack_tab= temp_wb.sheets["AR Rack"]
        rack_tab.api.Copy(After=wb.api.Sheets(2))
        rack_tab_it = wb.sheets[2]
        rack_tab_it.name = "Rack_Data(IT)"

        intial_date = rack_tab_it.range("B3").value.split("To")[0].strip()
        last_date = rack_tab_it.range("B3").value.split("To")[1].strip()

        intial_date_xl = f"01-01-{year}"

        last_date = f"{month}-{day}-{year}"
        xl_input_Date = intial_date_xl + f" To " + last_date
        rack_tab_it.range("B3").value = xl_input_Date

        rack_tab_it.activate()



        # bulk_tab_it = ""
        # delete_row_end = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
        rack_tab_it.api.Range(f"B9:J27").Delete(win32c.DeleteShiftDirection.xlShiftUp)


        input_tab.activate()
        input_tab.api.Range(f"{DD_No_column_letter}1").AutoFilter(Field:=f'{DD_No_column_no}', Criteria1:=['='])
        sp_lst_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"A2:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 
        input_tab.api.Range(f"A{sp_initial_rw}:A{sp_lst_row}").Copy(rack_tab_it.range(f"B100").api)


        rack_tab_it.activate()
        rack_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
        rack_tab_it.range(f"B9").api.EntireRow.Select()
        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

        
        ini = rack_tab_it.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
        rack_tab_it.range(f"B{ini}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
        ini_help = rack_tab_it.range(f'J'+ str(input_tab.cells.last_cell.row)).end('up').row
        ini = rack_tab_it.range(f'B{ini_help}').end('up').row

        rack_tab_it.api.Range(f"C8:I{ini}").Select()
        wb.app.api.Selection.FillDown()

        rack_tab_it.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        rack_tab_it.api.Range(f"B8:B{ini-1}").Font.Size = 9
        input_tab.activate()
        input_tab.api.AutoFilterMode=False
        input_tab.api.Range(f"{DD_No_column_letter}1").AutoFilter(Field:=f'{DD_No_column_no}', Criteria1:=['Total'])
        sp_lst_row = input_tab.range(f'H'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"H2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0] 

        input_tab.api.Range(f"H{sp_initial_rw}:H{sp_lst_row-1}").Copy(rack_tab_it.range(f"C8").api)
        input_tab.activate()
        
        input_tab.api.Range(f"I{sp_initial_rw}:M{sp_lst_row-1}").Copy(rack_tab_it.range(f"E8").api)

        rack_tab_it.range(f"E8:I{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        rack_tab_it.range(f"E8:I{ini-1}").api.Font.Size = 9
        rack_tab_it.range(f"C8:C{ini-1}").api.Font.Size = 9
        rack_tab_it.range(f"C8:C{ini-1}").number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'



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
        if int(sp_initial_rw)==1:
            pass
        else:
            bs_tab.range(f"I{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bs_tab.api.AutoFilterMode=False 
        time.sleep(1)
        bs_total = round(sum(bs_tab.range(f"J2").expand('down').value),2)
        bs_tab.range(f"I2").expand("table").copy(rack_tab_it.range(f"L8"))
        rack_tab_it.activate()
        rack_tab_it.autofit()
        bs_total_row = rack_tab_it.range(f'C{ini_help-1}').end('down').row
        rack_tab_it.range(f"C{bs_total_row}").value = bs_total
        #     Cells.Find(What:="accounts receivable", After:=ActiveCell, LookIn:= _
        # xlFormulas2, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        # xlNext, MatchCase:=False, SearchFormat:=False).Activate
        companny_name1 = rack_tab_it.range(f"B8:B{ini-1}").value
        refined_name1 = [" ".join(name.split(" ")[:-1]) for name in companny_name1]
        rack_tab_it.range(f"P8").options(transpose=True).value = refined_name1

        companny_name2= rack_tab_it.range(f"L8").expand('down').value
        refined_name2 = [name.strip() for name in companny_name2]
        rack_tab_it.range(f"L8").options(transpose=True).value = refined_name2

        rack_tab_it.range(f"J8").value = "=XLOOKUP(P8,L:L,M:M,0)"
        rack_tab_it.range(f"J8:J{ini-1}").api.Select()
                # bulk_tab_it.api.Range(f"C8:N{ini}").Select()
        wb.app.api.Selection.FillDown()
        rack_tab_it.range(f"J8").expand('down').number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        rack_tab_it.range(f"J8").expand('down').font.size = 9
        rack_tab_it.api.Range(f"J8:J{ini-1}").Copy()
        rack_tab_it.api.Range(f"J8:J{ini-1}")._PasteSpecial(Paste=-4163)
        wb.app.api.CutCopyMode=False
        rack_tab_it.range(f"L8").expand('table').delete()
        rack_tab_it.api.Range(f"N:N").EntireColumn.Delete()

        # bulk_tab_it.range(f"P8").expand("down").api.Copy(bulk_tab_it.range(f"L8").api)
        # bulk_tab_it.range(f"M8").expand('down').clear_contents()
        # bulk_tab_it.range(f"J8").expand("down").api.Copy(bulk_tab_it.range(f"M8").api)

        # bulk_tab_it.api.Range(f"P:P").EntireColumn.Delete()
        rack_tab_it.autofit()
        # bulk_tab2= temp_wb.sheets["Bulk(2)"]
        # bulk_tab2.api.Copy(After=wb.api.Sheets(3))
        # bulk_tab_it2 = wb.sheets[3]
        # bulk_tab_it2.name = "Bulk_Data(IT)(2)"

        # bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        # bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
        # brow_value = re.findall("\d+",bcell_value)[0]
        # bulk_tab_it2.range(f"B{int(brow_value)+1}").expand('table').delete()
        # bulk_tab_it2.range("B3").value = xl_input_Date

        # bulk_tab_it2.range(f"B9:J{int(brow_value)-1}").delete()

        # delete_row_end = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
        # delete_row_end2 = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
        # bulk_tab_it2.range(f"{delete_row_end2}:{delete_row_end2}").insert()
        # bulk_tab_it2.range(f"{delete_row_end2+1}:{delete_row_end+1}").delete()


        # bulk_tab_it.api.Range(f"B8:C{ini-1}").Copy(bulk_tab_it2.range(f"B100").api)


        # bulk_tab_it2.activate()
        # bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Copy()
        # bulk_tab_it2.range(f"B9").api.EntireRow.Select()
        # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

        # ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row
        # bulk_tab_it2.range(f"B{ini2}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
        # ini2 = bulk_tab_it2.range(f'B'+ str(input_tab.cells.last_cell.row)).end('up').end('up').row

        # bulk_tab_it2.api.Range(f"D8:J{ini2-1}").Select()
        # wb.app.api.Selection.FillDown()

        # bulk_tab_it2.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)

        # bulk_tab_it2.api.Range(f"B{ini2-1}").Font.Bold = True

        # bulk_tab_it.api.Range(f"E8:I{ini-1}").Copy(bulk_tab_it2.range(f"E8").api)

        rack_tab_it.api.Range(f"J1").Copy()
        rack_tab_it.api.Range(f"C8:C{ini-1}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        rack_tab_it.api.Range(f"E8:I{ini-1}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        wb.app.api.CutCopyMode=False

        # bs_total_row2 = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').row
        # bulk_tab_it2.range(f"C{bs_total_row2}").value = -bs_total
        companny_name = rack_tab_it.range(f"B8:B{ini-1}").value
        refined_name = [" ".join(name.split(" ")[:-1]) + " " for name in companny_name]
        rack_tab_it.range(f"B8").options(transpose=True).value = refined_name

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
        rack_tab_it.activate()
        del_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
        rack_tab_it.range(f'B{del_row}').expand('table').delete()
        rack_tab_it.api.Range(f"L8").Value="=+XLOOKUP(B8,'[Group_mapping.xlsx]Sheet1'!$A:$A,'[Group_mapping.xlsx]Sheet1'!$B:$B,0)"

        rack_tab_it.api.Range(f"L8:L{ini-1}").Select()
        wb.app.api.Selection.FillDown()
        rack_tab_it.api.Range(f"L7").Select()
        rack_tab_it.api.Range(f"L6").Value = "Xlookup"
        rack_tab_it.api.AutoFilterMode=False
        rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:='=0')
        
        sp_lst_row = rack_tab_it.range(f'L'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
        if sp_lst_row != 8:
            sp_address= rack_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        else:
            sp_initial_rw = 8

        rack_tab_it.range(f"L{sp_initial_rw}").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
        try:
            rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1)
        except:
            pass    
        font_colour,Interior_colour = conditional_formatting(range=f"L:L",working_sheet=rack_tab_it,working_workbook=wb)

        rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        sp_lst_row = rack_tab_it.range(f'L'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
        sp_address= rack_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]

        try:
            rack_tab_it.range(f"L{sp_initial_rw}:L{sp_lst_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(rack_tab_it.range(f"B100").api)
        except:
            pass  

        if rack_tab_it.range(f"B100").expand('down').value !=None:
            grp_cm_list = rack_tab_it.range(f"B100").expand('down').value
        # bulk_tab_it2.range(f"B100").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            grp_cm_list2 = list(set(grp_cm_list))
            rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
            val_row = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
            if len(grp_cm_list2)>0:
                for i in range(len(grp_cm_list2)):
                    # if i >0:
                    #     val_row = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row-2
                    rack_tab_it.api.Range(f"L7").Select()
                    rack_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=[grp_cm_list2[i]])
                    sp_lst_row = rack_tab_it.range(f'L'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
                    sp_address= rack_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
                    sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
                    if rack_tab_it.range(f"C{sp_initial_rw}").value + rack_tab_it.range(f"C{sp_lst_row}").value<0:
                        # in_rw = bulk_tab_it2.range(f'B'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').row
                        rack_tab_it.range(f"{sp_initial_rw}:{sp_lst_row}").api.EntireRow.Copy()
                        # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                        rack_tab_it.range(f"{val_row+2}:{val_row+2}").api.EntireRow.Select()
                        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                        rack_tab_it.range(f"{sp_initial_rw}:{sp_lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                    else:
                        print("second case")

                rack_tab_it.api.Cells.FormatConditions.Delete()
                rack_tab_it.api.AutoFilterMode=False

        rack_tab_it.api.Range(f"L:L").EntireColumn.Delete()
        rack_tab_it.activate()
        font_colour,Interior_colour = conditional_formatting2(range=f"C8:C{ini-1}",working_sheet=rack_tab_it,working_workbook=wb)
        rack_tab_it.api.Range(f"C7").AutoFilter(Field:=2, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

        sp_lst_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
        sp_address= rack_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        if int(sp_initial_rw)==6:
            rack_tab_it.api.Range(f"C7").AutoFilter(Field:=2)
        elif int(sp_lst_row) ==int(sp_initial_rw):
            rack_tab_it.range(f"B{sp_initial_rw}").expand("right").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(rack_tab_it.range(f"B100").api)
            count_n = rack_tab_it.range(f"B{sp_initial_rw}").expand("right").api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Count
        else:    
            rack_tab_it.range(f"B{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(rack_tab_it.range(f"B100").api)


        # value_row = bulk_tab_it2.range(f'C'+ str(bulk_tab_it2.cells.last_cell.row)).end('up').end('up').end('up').row
        if rack_tab_it.range(f"B100").value!=None:
            rack_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
            try:
                rack_tab_it.range(f"A{val_row+2}").api.EntireRow.Select()
            except:
                val_row = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
                rack_tab_it.range(f"A{val_row+2}").api.EntireRow.Select()
            wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
            wb.app.api.CutCopyMode=False


            if count_n!=1:
                rw_faltu=rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row  
                rack_tab_it.range(f"B{rw_faltu}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            else:
                rw_faltu=rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row      
                rack_tab_it.range(f"B{rw_faltu}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)



        if int(sp_initial_rw)==6:
            pass
        elif int(sp_lst_row) ==int(sp_initial_rw):
            rack_tab_it.range(f"B{sp_initial_rw}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        else:    
            rack_tab_it.range(f"B{sp_initial_rw}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        rack_tab_it.api.AutoFilterMode=False

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
        company_sheet.range(f"A2").expand('down').api.Copy(rack_tab_it.range(f"B100").api)
        rack_tab_it.api.Cells.FormatConditions.Delete()
        rack_tab_it.activate()
        font_colour,Interior_colour = conditional_formatting(range=f"B:B",working_sheet=rack_tab_it,working_workbook=wb)
        rack_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

        sp_lst_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
        sp_address= rack_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
        if rack_tab_it.api.Range(f"B{sp_initial_rw}").Value==None:
            pass
        else:
            print("please check for code this condition is new")
            value_row2 = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').end('up').row

            rack_tab_it.range(f"B{sp_initial_rw}").expand('table').api.Copy(rack_tab_it.range(f"B150").api)

            rack_tab_it.range(f"B150").expand('table').api.EntireRow.Copy()
            # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
            rack_tab_it.range(f"A{value_row2+1}").api.EntireRow.Select()
            wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
            rack_tab_it.range(f"B{sp_initial_rw}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            rack_tab_it.api.AutoFilterMode=False
            rack_tab_it.api.Cells.FormatConditions.Delete()

        # faltu_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
        # rack_tab_it.range(f"b{faltu_row}").expand('table').delete()

        input_tab.api.AutoFilterMode=False
        rack_tab_it.api.AutoFilterMode=False
        rack_tab_it.api.Cells.FormatConditions.Delete()
        
        faltu_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
        rack_tab_it.range(f"b{faltu_row}").expand('down').delete()

        t_df.fillna(0,inplace= True)
        t_df = t_df[t_df.COMPANY.isin(company_names) == False]
        grp_df = t_df.groupby(['COMPANY'], sort=False)['Balance','< 10','11 - 30','31 - 60','61 - 90','> 90'].sum().reset_index()
        grp_df.insert(2,"> 10",grp_df[['11 - 30','31 - 60','61 - 90','> 90']].sum(axis=1))
        grp_df['As Per BS'] = grp_df['Balance'] - grp_df['< 10'] - grp_df['> 10']
        for i in range(len(grp_df['COMPANY'])):
            grp_df['COMPANY'][i] = " ".join(grp_df['COMPANY'][i].split(" ")[:-1]) + f" "

        # bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        # bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
        check_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
        if rack_tab_it.range(f"B{check_row}").value=='Total':
            brow_value = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').row + 2
        else:
            brow_value = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row + 1 

        rack_tab_it.api.Range(f"B{int(ini)}:B{int(ini)+len(grp_df)-1}").EntireRow.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
        rack_tab_it.range(f'B{int(ini)}').options(index = False,header=False).value = grp_df 

        # rack_tab_it.range(f'B{int(brow_value)}').expand('down').font.bold= False


        rack_tab_it.range(f"B8:J{ini-1}").api.Sort(Key1=rack_tab_it.range(f"B8:B{ini-1}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
      
        rack_tab_it.range(f'B{int(ini)}').expand('table').api.Sort(Key1=rack_tab_it.range(f'B{int(ini)+1}').expand('down').api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        tell_row = rack_tab_it.range(f'B{int(brow_value)}').end('down').row 
        count = 0 + count_n
        for i in range(len(grp_df['COMPANY'])):
            conditional_formatting(range=f'B8:B{tell_row}',working_sheet=rack_tab_it,working_workbook=wb)
            rack_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
            rack_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=[grp_df['COMPANY'][i]])
            sp_lst_row = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
            sp_address= rack_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
            int_check = rack_tab_it.range(f"B{sp_initial_rw}").expand("table").get_address().split(":")[-1]
            lst_row = re.findall("\d+",int_check .replace("$","").split(":")[0])[0]
            if rack_tab_it.range(f"C{sp_initial_rw}").value + rack_tab_it.range(f"C{lst_row}").value<=1:
                rack_tab_it.range(f"{lst_row}:{lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                in_rw = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
                if count>=1:
                    in_rw = rack_tab_it.range(f'B'+ str(rack_tab_it.cells.last_cell.row)).end('up').row
                    rack_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.EntireRow.Copy()
                    # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                    rack_tab_it.range(f"{in_rw+1}:{in_rw+1}").api.EntireRow.Select()
                    wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                else:    
                    rack_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.EntireRow.Copy()
                    # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                    rack_tab_it.range(f"{in_rw+2}:{in_rw+2}").api.EntireRow.Select()
                    wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                rack_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                rack_tab_it.api.AutoFilterMode=False
                rack_tab_it.api.Cells.FormatConditions.Delete()
                count+=1
            else:
                print("second case")
                rack_tab_it.api.AutoFilterMode=False
                rack_tab_it.api.Cells.FormatConditions.Delete()

        #ineligible accounts check
        # bulk_tab_it2.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it2.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        # bcell_value = bulk_tab_it2.api.Application.ActiveCell.Address.replace("$","")
        # brow_value = re.findall("\d+",bcell_value)[0]
       
        # if bulk_tab_it2.range(f"B{int(brow_value)+1}").value!=None:
        #     pass
        # else:
        #     bulk_tab_it2.range(f"{brow_value}:{brow_value}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)


        #updating formula

        rack_tab_it.api.Cells.Find(What:="Total", After:=rack_tab_it.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = rack_tab_it.api.Application.ActiveCell.Address.replace("$","")
        brow_value = re.findall("\d+",bcell_value)[0]

        formula_row = int(brow_value)+4

        pre_row = rack_tab_it.range(f"C{formula_row}").end('up').row

        fst_rng = rack_tab_it.range(f"C8").expand("down").get_address().replace("$","")

        if type(rack_tab_it.range(f"C{formula_row}").end("down").value)==float:

            rw = rack_tab_it.range(f'C'+ str(rack_tab_it.cells.last_cell.row)).end('up').end('up').row
            mid_range = rack_tab_it.range(f"C{rw}").expand("down").get_address().replace("$","")
            rack_tab_it.range(f"C{formula_row}").formula = f"=+C{pre_row}-SUM({fst_rng})-SUM({mid_range})"
        else:
            rack_tab_it.range(f"C{formula_row}").formula = f"=+C{pre_row}-SUM({fst_rng})" 

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

        tablist={input_tab:win32c.ThemeColor.xlThemeColorAccent2,rack_tab_it:win32c.ThemeColor.xlThemeColorAccent6}
        for tab,color in tablist.items():
                tab.activate()
                tab.api.Tab.ThemeColor = color
                tab.autofit()
                tab.range(f"A1").select()
        initial_tab.activate()
        initial_tab.range(f"A1").select()
        wb.save(f"{output_location}\\AR Aging Rack {month}{day}-updated"+'.xlsx') 
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