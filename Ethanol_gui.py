from email.mime import message
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
 

def purchased_ar(input_date, output_date):
    try:       
        job_name = 'purchased_ar_automation'
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
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
        input_tab.api.Range(f"A1:M{lsr_rw}").AutoFilter(Field:=4, Criteria1:=[f'>={datetime.now().date().replace(day=int(day),month=int(month))}'])

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
    wp_job_ids = {'ABS':1,'Purchased AR Report':purchased_ar}
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

