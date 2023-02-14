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





def last_day_of_month(date):
    if date.month == 12:
        return date.replace(day=31)
    return date.replace(month=date.month+1, day=1) - timedelta(days=1)





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
        working_workbook.activate()
        working_sheet.activate()
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
def conditional_formatting_uniq(range:str,working_sheet,working_workbook):
    try:
        working_workbook.activate()
        working_sheet.activate()
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
        working_workbook.activate()
        working_sheet.activate()
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


def knockOffAmtDiff(curr,final, wb, input_sht, input_sht2, credit_col_letter, debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter, row_dict, eth_trueup_col_letter=None):
    try:
        print(row_dict["Knock_Off"])
        if abs(input_sht.range(f"{credit_col_letter}{curr}").value) == abs(input_sht2.range(f"{debit_col_letter}{final}").value):
            print(f"Moving {curr} to knockoff")
            knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row

            #Copy Pasting Whole data
            # input_sht.range(f"{curr}:{final}").api.Copy()
            # wb.activate()
            # knock_off_sht.activate()
            # knock_off_sht.range(f"A{knock_off_last_row+1}").api.Select()
            # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
            # knock_off_sht.autofit()
            if input_sht==input_sht2:
                # input_sht.range(f"{curr}:{final}").copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))

                # input_sht.range(f"{curr}:{final}").delete()
                # input_sht.range(f"{curr}:{final}").color ="#00FF00"
                
                if not len(row_dict["Knock_Off"]):
                    row_dict["Knock_Off"] = [[f"{curr}:{final}"]]
                    # knockoff_list.append(f"{curr}:{final}")
                # elif int(knockoff_list[-1].split(":")[-1]) == curr-1:   #prev final == currnt -1
                elif len(row_dict["Knock_Off"][-1]) <=24:
                    if int(row_dict["Knock_Off"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1
                        # knockoff_list[-1] = f'{knockoff_list[-1].split(":")[0]}:{final}'
                        row_dict["Knock_Off"][-1][-1] = f'{row_dict["Knock_Off"][-1][-1].split(":")[0]}:{final}'
                    else:
                        # knockoff_list.append(f"{curr}:{final}")
                        row_dict["Knock_Off"][-1].append(f"{curr}:{final}")
                elif len(row_dict["Knock_Off"][-1]) >24:
                    row_dict["Knock_Off"].append([f"{curr}:{final}"])
                
            else:
                input_sht.range(f"{curr}:{curr}").copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))
                input_sht2.range(f"B{final}:{eth_trueup_col_letter}{final}").copy(knock_off_sht.range(f"A{knock_off_last_row+2}"))

                #shifting credit amount to right cell copied from ethanol accrual
                knock_off_sht.range(f"K{knock_off_last_row+2}").copy(knock_off_sht.range(f"L{knock_off_last_row+2}"))
                knock_off_sht.range(f"K{knock_off_last_row+2}").clear()
                knock_off_sht.range(f"M{knock_off_last_row+2}").clear()#Clearing Final Amount

                input_sht.range(f"{curr}:{curr}").delete()
                input_sht2.range(f"{final}:{final}").delete()
                curr-=1
                

                # input_sht.range(f"{curr}:{curr}").delete()
                # input_sht.range(f"{curr}:{curr}").color ="#00FF00"
                # input_sht2.range(f"{final}:{final}").delete()
                # input_sht.range(f"{final}:{final}").color ="#00FF00"
                # if not len(row_dict["Knock_Off"]):
                #     row_dict["Knock_Off"] = [[f"{curr}:{final}"]]                   
                # elif len(row_dict["Knock_Off"][-1]) <=24:
                #     if int(row_dict["Knock_Off"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1                       
                #         row_dict["Knock_Off"][-1][-1] = f'{row_dict["Knock_Off"][-1][-1].split(":")[0]}:{final}'
                #     else:
                #         row_dict["Knock_Off"][-1].append(f"{curr}:{final}")
                # elif len(row_dict["Knock_Off"][-1]) >24:
                    # row_dict["Knock_Off"].append([f"{curr}:{final}"])
            # curr-=1
        elif (abs(input_sht.range(f"{credit_col_letter}{curr}").value) - abs(input_sht2.range(f"{debit_col_letter}{final}").value))!=0:
            #amt diff
            print(f"Moving {curr} to amount diff")
            amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row

            if input_sht==input_sht2:
                pass
                # input_sht.range(f"{curr}:{final}").api.Copy()
                # wb.activate()
                # amt_diff_sht.activate()
                # amt_diff_sht.range(f"A{amt_diff_last_row+1}").api.Select()
                # wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                # amt_diff_sht.autofit()
                # input_sht.range(f"{i}:{i+1}").copy(amt_diff_sht.range(f"A{amt_diff_last_row+1}"))

                # input_sht.range(f"{curr}:{final}").delete()
                # input_sht.range(f"{curr}:{final}").color ="#FFFF00"
                if not len(row_dict["Amt_Dff"]):
                    row_dict["Amt_Dff"] = [[f"{curr}:{final}"]]                   
                elif len(row_dict["Amt_Dff"][-1]) <=24:
                    if int(row_dict["Amt_Dff"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1                       
                        row_dict["Amt_Dff"][-1][-1] = f'{row_dict["Amt_Dff"][-1][-1].split(":")[0]}:{final}'
                    else:
                        row_dict["Amt_Dff"][-1].append(f"{curr}:{final}")
                elif len(row_dict["Amt_Dff"][-1]) >24:
                    row_dict["Amt_Dff"].append([f"{curr}:{final}"])
            else:
                input_sht.range(f"{curr}:{curr}").api.Copy()
                wb.activate()
                amt_diff_sht.activate()
                amt_diff_sht.range(f"A{amt_diff_last_row+1}").api.Select()
                wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                
                input_sht2.range(f"B{final}:{eth_trueup_col_letter}{final}").api.Copy()
                wb.activate()
                amt_diff_sht.activate()
                amt_diff_sht.range(f"A{amt_diff_last_row+2}").api.Select()
                wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

                amt_diff_sht.range(f"K{knock_off_last_row+2}").copy(amt_diff_sht.range(f"L{knock_off_last_row+2}"))
                amt_diff_sht.range(f"K{knock_off_last_row+2}").clear()
                amt_diff_sht.range(f"M{knock_off_last_row+2}").clear()#Clearing Final Amount
                

                amt_diff_sht.autofit()
                # input_sht.range(f"{i}:{i+1}").copy(amt_diff_sht.range(f"A{amt_diff_last_row+1}"))

                input_sht.range(f"{curr}:{curr}").delete()
                # input_sht.range(f"{curr}:{curr}").color ="#FFFF00"
                input_sht2.range(f"{final}:{final}").delete()
                # input_sht.range(f"{final}:{final}").color ="#FFFF00"
                curr-=1

                if not len(row_dict["Amt_Dff"]):
                    row_dict["Amt_Dff"] = [[f"{curr}:{final}"]]                   
                elif len(row_dict["Amt_Dff"][-1]) <=24:
                    if int(row_dict["Amt_Dff"][-1][-1].split(":")[-1]) == curr-1:   #prev final == currnt -1                       
                        row_dict["Amt_Dff"][-1][-1] = f'{row_dict["Amt_Dff"][-1][-1].split(":")[0]}:{final}'
                    else:
                        row_dict["Amt_Dff"][-1].append(f"{curr}:{final}")
                elif len(row_dict["Amt_Dff"][-1]) >24:
                    row_dict["Amt_Dff"].append([f"{curr}:{final}"])

            # curr-=1
        else:
            #line for ethnaol accrual tab
            print(f'current line {curr} remains here for ethanol accrual tab having mrn no.{input_sht.range(f"{mrn_col_letter}{curr}")}')
        return curr, row_dict
    except Exception as e:
        raise e



def row_range_calc(filter_col:str, input_sht,wb):
    sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row

    sp_address= input_sht.api.Range(f"{filter_col}2:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address

    sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

    row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

    while row_range[-1]!=sp_lst_row:

        sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row

        sp_address.extend(input_sht.api.Range(f"{filter_col}{row_range[-1]}:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)

        # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

        # row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
        
    
    sp_address = sp_address.replace("$","").split(",")
    init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
    sublist = []
    flat_list = [item for sublist in init_list for item in sublist]
    return flat_list, sp_lst_row,sp_address



def thick_bottom_border(cellrange:str,working_sheet,working_workbook):
        working_sheet.api.Range(cellrange).Select()
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeLeft).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeTop).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlEdgeRight).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideVertical).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlInsideHorizontal).LineStyle = win32c.Constants.xlNone
        linestylevalues=[win32c.BordersIndex.xlEdgeBottom]
        for values in linestylevalues:
            a=working_workbook.app.selection.api.Borders(values)
            a.LineStyle = win32c.LineStyle.xlContinuous
            a.ColorIndex = 0
            a.TintAndShade = 0
            a.Weight = win32c.BorderWeight.xlMedium


def mrn_pdf_extractor(pdf_file, mrn_dict, date_list, rack=False):
    try:
        if not rack:
            df = read_pdf(pdf_file , pages = 'all', guess = False, stream = True ,
                pandas_options={'header':None}, area = ["110,40,530,665"], columns = ["70,125,200,250,280,365,415,460,560,610,660"])
            df = pd.concat(df, ignore_index=True)
        else:
            df = read_pdf(pdf_file , pages = 'all', guess = False, stream = True ,
                pandas_options={'header':None}, area=['110,30,560,680'], columns = ["50,125,200,250,280,365,415,460,560,610,630"])
            df = pd.concat(df, ignore_index=True)
        print(df)
        
        for row in range(len(df)):
            if df[1][row] == "AL: LOCATION" and not rack:
                try:
                    pdf_date = datetime.strptime(df[6][row-1], "%m/%d/%Y")#Picking tikcet open dates
                    # mrn_dict[df[1][row-1]].append([df[9][row], df[0][row-1], df[6][row-1]])
                    mrn_dict[int(df[0][row-1])].append([pdf_date, df[9][row]])
                    if pdf_date not in date_list:
                        if pdf_date.day<16:
                            date_list.append(pdf_date)
                except:
                    pdf_date = datetime.strptime(df[6][row-1], "%m/%d/%Y")
                    # mrn_dict[df[1][row-1]] = [[df[9][row], df[0][row-1], df[6][row-1]]]
                    mrn_dict[int(df[0][row-1])] = [[pdf_date, df[9][row]]]
                    if pdf_date not in date_list:
                        if pdf_date.day<16:
                            date_list.append(pdf_date)
            elif df[1][row] == "AL: LOCATION"   and rack:
                try:
                    mrn_dict[df[1][row-1]] = int(df[11][row].replace(',',''))
                except:
                    try:
                        mrn_dict[df[1][row-1]] = df[11][row]
                    except Exception as e:
                        raise e
        return date_list, mrn_dict
    except Exception as e:
        raise e


def rack_pdf_data_extractor(pdf_loc):
    try:
        # date_area=["8.798,105.876,47.048,508.266"]
        df=read_pdf(pdf_loc,stream=True, multiple_tables=True,pages='all',silent=True,guess=False)
        df = read_pdf(pdf_loc , pages = 'all', guess = False, stream = True ,
                pandas_options={'header':None}, area = ["100,50,565,770"], columns = ["115,140,170,228,290,355,420,482,546,595,616,670,690,750"])
        dates = read_pdf(pdf_loc , pages = 'all', guess = False, stream = True ,
                    pandas_options={'header':None}, area = ["40,400,46,430"])
        file_date = (datetime.strptime(dates[1][0][0], "%m/%d/%y")).strftime("%Y%m%d")
        df = pd.concat(df, ignore_index=True)
        #removing row from where Total starts
        last_row = df.loc[df[0]=='TOTAL'].index[0]
        df = df[:last_row]
        return df, file_date
    except Exception as e:
        raise e

# def common():
#     try:
#         set_borders()
#         freezepanes_for_tab()
#         interior_coloring()
#         conditional_formatting2()
#         interior_coloring_by_theme()
#         num_to_col_letters()
#         insert_all_borders()
#         conditional_formatting()
#         knockOffAmtDiff()
#         row_range_calc()
#         thick_bottom_border()
        
        



#     except Exception as e:
#         raise e
