import logging
import os, sys
import time
from datetime import datetime
import numpy as np
# import bu_alerts
from datetime import datetime

import xlwings as xw
# from tabula import read_pdf
# import openpyxl as xl
from xlwings.constants import DeleteShiftDirection
import xlwings.constants as win32c
import random
# import webcolors
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import re


JOBNAME = "FOB_BBR_PROCESS"
# FILE ="C:\DEEPFOLDER\Tasks\BBR_PROCESS\BBR_20221130\MRN 2022.xlsx"
# FILE2 = "C:\DEEPFOLDER\Tasks\BBR_PROCESS\BBR_20221130\FOB Inventory_20221130.xlsx"

# wb = xw.Book(FILE)
# wb1 = xw.Book(FILE2)

# adte = datetime.date.today()
# B/L date

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
    



def row_range_calc(filter_col:str, ws1):
    sp_lst_row = ws1.range(f'{filter_col}'+ str(ws1.cells.last_cell.row)).end('up').row

    sp_address= ws1.api.Range(f"{filter_col}2:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address

    sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

    row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

    while row_range[-1]!=sp_lst_row:

        sp_lst_row = ws1.range(f'{filter_col}'+ str(ws1.cells.last_cell.row)).end('up').row

        sp_address.extend(ws1.api.Range(f"{filter_col}{row_range[-1]}:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)

        # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

        # row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
        
    
    sp_address = sp_address.replace("$","").split(",")
    init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
    sublist = []
    flat_list = [item for sublist in init_list for item in sublist]
    return flat_list, sp_lst_row,sp_address





def fob(wb,wb1,wb2,ws1,ws2,ws3):

    try:
        print('hello')
        ws = wb.sheets("Sheet1")
        wa = ws.copy(name = "SheetA", after = wb.sheets["Sheet1"])
        wa.range('A:A').delete()
        wa.range('1:5').api.Delete(DeleteShiftDirection.xlShiftUp)
        wb.activate()
        wa.activate()
        
        
        curr_col_list = wa.range("A1").expand('right').value
        bl_date_col = curr_col_list.index("Arrival Date")
        bl_date_col_letters = num_to_col_letters(bl_date_col+1)
        last_row = wa.range(f'A'+ str(wa.cells.last_cell.row)).end('up').row
        wa.api.AutoFilterMode = False 
        # wa.api.Range(f"{bl_date_col_letters}1:{bl_date_col_letters}{last_row}").AutoFilter(Field:=f"{bl_date_col+1}", Criteria1="<>")
        wa.api.Range(f"A1:{bl_date_col_letters}{last_row}").AutoFilter(Field:=f"{bl_date_col+1}", Criteria1="<>")
        # wa.used_range.api.AutoFilter(Field:=17)
        # wa.range(str('A1') + ':65536').api.Delete(DeleteShiftDirection.xlShiftUp)
        wa.api.Range(f"A2:{bl_date_col_letters}{last_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.selection.delete(shift='left')
        wa.api.AutoFilterMode = False


        vendor_ref_col = curr_col_list.index("Vendor Ref")
        vendor_ref_col_letters = num_to_col_letters(vendor_ref_col+1)
        last_row = wa.range(f'A'+ str(wa.cells.last_cell.row)).end('up').row
        wa.api.Range(f"A1:{vendor_ref_col_letters}{last_row}").AutoFilter(Field:=f"{vendor_ref_col+1}", Criteria1="Total")
        wa.api.Range(f"A2:{vendor_ref_col_letters}{last_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.selection.delete(shift='left')
        wa.api.AutoFilterMode = False


        code_col = curr_col_list.index("Code")
        code_col_letters = num_to_col_letters(code_col+1)
        last_row = wa.range(f'A'+ str(wa.cells.last_cell.row)).end('up').row
        wa.api.Range(f"A1:{code_col_letters}{last_row}").AutoFilter(Field:=f"{code_col+1}", Criteria1="=")
        wa.api.Range(f"A2:{code_col_letters}{last_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.selection.delete(shift='left')
        wa.api.AutoFilterMode = False


        inco_terms_col = curr_col_list.index("Inco terms")
        inco_terms_col_letters = num_to_col_letters(inco_terms_col+1)
        last_row = wa.range(f'A'+ str(wa.cells.last_cell.row)).end('up').row
        wa.api.Range(f"A1:{inco_terms_col_letters}{last_row}").AutoFilter(Field:=f"{inco_terms_col+1}", Criteria1="DEL")
        wa.api.Range(f"A2:{inco_terms_col_letters}{last_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.selection.delete(shift='left')
        wa.api.AutoFilterMode = False


        try:
           terminal_col = curr_col_list.index("Terminal")
           terminal_col_letters = num_to_col_letters(terminal_col+1)
           last_row = wa.range(f'A'+ str(wa.cells.last_cell.row)).end('up').row
           wa.api.Range(f"A1:{terminal_col_letters}{last_row}").AutoFilter(Field:=f"{terminal_col+1}", Criteria1="<>*tload",criteria2="<>*Tank")
           wa.api.Range(f"A2:{terminal_col_letters}{last_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
           wb.app.selection.delete(shift='left')
           wa.api.AutoFilterMode = False
        except:
            pass

        try:
            wb.sheets['SheetA'].range('A1:AH1').expand('down').copy()
            ws1 =wb1.sheets['FOB report_20221130'] 
            ws1.activate()
            # we = wb1.sheets.add('Test', after='FOB report_20221130')
            ws1.range('B2').paste()
            ws1.range('A2').expand('right').delete()
            curr_col_list2 = ws1.range("B1").expand('right').value
            voucher_no_col = curr_col_list2.index("Voucher No.-Base ")
            voucher_no_col_letters = num_to_col_letters(voucher_no_col+2)
            last_row = ws1.range(f'B'+ str(wa.cells.last_cell.row)).end('up').row
            YELLOW = win32c.RgbColor.rgbYellow
            ws1.api.Range(f"B2:{voucher_no_col_letters}{last_row}").AutoFilter(Field:=f"{voucher_no_col+1}",Criteria1=YELLOW,Operator = win32c.AutoFilterOperator.xlFilterCellColor)
            
            flat_list, sp_lst_row,sp_address = row_range_calc("B",ws1)
            print(sp_address)
            wb1.activate()
            ws1.activate()
            ws1.range(sp_address[0]).select()
            for address in sp_address:
                        address = voucher_no_col_letters + (f":AI").join(address.split(":"))
                        print(address)

            ws1.range(address).clear()#api.Delete(DeleteShiftDirection.xlShiftToLeft)
            ws1.api.AutoFilterMode = False
            YELLOW = win32c.RgbColor.rgbYellow
            ws2 = wb1.sheets['Inv Back Track']
            ws2.activate()
            last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            ws2.api.Range(f"A1:A{last_row}").AutoFilter(Fileld:=1, Criteria1:=YELLOW, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
            ws2.range(f'A2:AI2').expand('down').copy()
            last_row0 = ws1.range(f'B'+ str(wa.cells.last_cell.row)).end('up').row
            ws1.range(f'B{last_row0+1}').paste()

            # now MRN_November 2022
            frt_amount = curr_col_list2.index("FrtAmt")
            frt_amount_letters = num_to_col_letters(frt_amount+2)
            last_row_frt = ws1.range(f'AK'+ str(ws1.cells.last_cell.row)).end('up').row
            ws3.activate()
            # = +XLOOKUP(G2,'[MRN_November2022.xlsx]MRN_November2022'!$F:$F,'[MRN_November2022.xlsx]MRN_November2022'!$AU:$AX,0)
            # ws1.range(f"{frt_amount_letters}{last_row_frt}").formula = f"=+XLOOKUP({ws2.range('G2')},'[{ws3}]{ws3.range('F')},'[{ws3}]{ws3.range('AU:AX')}"
            ws1.range(f"AK2:AK{last_row_frt}").formula = f"=+XLOOKUP(G2,'[{wb2.name}]{ws3.name}'!$F:$F,'[{wb2.name}]{ws3.name}'!$AU:$AX,0)"

            print("DONE")

        except Exception as e:
            raise e

    except Exception as e:
        raise e


# if __name__ == "__main__":
def fob_runner(start_date,end_date):
    try:
        start_date2 = datetime.strftime(datetime.strptime(start_date,"%m.%d.%Y"), "%Y%m%d")
        start_date3 = datetime.strftime(datetime.strptime(start_date,"%m.%d.%Y"), "%Y")
        global FILE
        FILE =  f"J:\India\BBR\IT_BBR\Reports\FOB\MRN {start_date3}.xlsx"
        FILE2 = f"J:\India\BBR\IT_BBR\Reports\FOB\FOB Inventory_{start_date2}.xlsx"
        FILE3 = f"J:\India\BBR\IT_BBR\Reports\FOB\MRN_November {start_date3}.xlsx"
        wb = xw.Book(FILE)
        wb1 = xw.Book(FILE2)
        ws1 =wb1.sheets[f'FOB report_{start_date2}'] 
        global ws2
        ws2 = wb1.sheets['Inv Back Track']
        wb2 = xw.Book(FILE3,update_links=False)
        ws3 = wb2.sheets[f'MRN_November {start_date3}']

        fob(wb,wb1,wb2,ws1,ws3,ws2)

    except Exception as e:
        raise e

