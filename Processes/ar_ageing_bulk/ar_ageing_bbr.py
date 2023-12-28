import re
import os
import logging
import glob, time
import numpy as np
import pandas as pd
import xlwings as xw
from tkinter import messagebox
import xlwings.constants as win32c
from datetime import date, timedelta
# from common import set_borders,freezepanes_for_tab,interior_coloring,conditional_formatting2,interior_coloring_by_theme,num_to_col_letters,insert_all_borders,conditional_formatting,knockOffAmtDiff,row_range_calc,thick_bottom_border
from Common.common import set_borders,freezepanes_for_tab,interior_coloring,conditional_formatting2,interior_coloring_by_theme,num_to_col_letters,insert_all_borders,conditional_formatting,knockOffAmtDiff,row_range_calc,thick_bottom_border

def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        print(f"Exception caught in num_to_col_letters method: {e}")
        logging.info(f"Exception caught in num_to_col_letters method: {e}")
        raise e
    
    
def xlOpner(inputFile):
    try:
        retry = 0
        while retry<10:
            try:
                input_wb = xw.Book(inputFile, update_links=False)
                return input_wb
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
    except Exception as e:
        print(f"Exception caught in xlOpner :{e}")
        logging.info(f"Exception caught in xlOpner :{e}")
        raise e
    
    
def range_divider(list1, list2,start_col,end_col):
    try:
        modified_list1 = []

        for range_str in list1:
            start, end = range_str.split(':')
            start = int(start)
            end = int(end)
            modified_range = []
            for num in range(start, end+1):
                if num not in list2:
                    modified_range.append(num)
            
            if modified_range:
                modified_list1.append(f"{start_col}{min(modified_range)}:{end_col}{max(modified_range)}")

        print(modified_list1)
        return modified_list1
    except Exception as e:
        print(f"Exception caught in range_divider method: {e}")
        logging.info(f"Exception caught in range_divider method: {e}")
        raise e
    
    
def row_range_calc(filter_col:str, input_sht,wb):
    try:
        sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row
        sp_address= input_sht.api.Range(f"{filter_col}2:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]

        row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

        while row_range[-1]!=sp_lst_row:
            sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row
            sp_address = sp_address+','+(input_sht.api.Range(f"{filter_col}{row_range[-1]+1}:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)
            row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
        sp_address = sp_address.replace("$","").split(",")
        init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
        sublist = []
        flat_list = [item for sublist in init_list for item in sublist]
        return flat_list, sp_lst_row,sp_address
    except Exception as e:
        print(f"Exception caught in row_range_calc method: {e}")
        logging.info(f"Exception caught in row_range_calc method: {e}")
        raise e


def get_col_list(inp_sht):
    try:
        col_list = inp_sht.range('A1:W1').value
        last_col_letter=num_to_col_letters(len(col_list))
        return col_list
    except Exception as e:
        print(f"Exception caught in get_col_list method: {e}")
        logging.info(f"Exception caught in get_col_list method: {e}")
        raise e

#Done---------------------------

# 1.Remove Unnecessary Columns
# 2.Apply filter on Future reminist and 0-10 column
# 3.Merge Future  Reminist and 0-10 column
# 4.Fill blank_cells
# 5.Copy the data into dataframe
# 6.Sort Customer name and Due date column 
    # 6.1 First sort by Due Date  --Oldest To Newest 
    # 6.2 Sort By Customer Name
# 7.Group by Customername and sum the other columns

# before pasting make a seperate dataframe of demurrage, friet factoring and cis concerns.
# 8.Copy template in other sheet
# 9.Paste df into this sheet retaining formulas
# 10.Get Cis concern and non products from the df and paste them in lower part after multiplying with -1
# 11.If total for this cis concerns =0 or negative, cut this from here and place in other table below this table
# 12.Verify sum from balance sum obtained from balance sheet(manually obtained)
# 13.get names from mapping sheet using vlook up.


def ar_ageing_BBR(input_date,output_date):
    try:
        
        job_name = 'ar_ageing_Bulk'
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
        year = input_date.split(".")[-1]
        
        # template_loc=os.getcwd()+r"\Other Files\Templates\AR Aging Bulk Template.xlsx"
        # grp_sheet_loc=os.path.join(os.getcwd()+r"\Other Files\Template_File\Group_mapping.xlsx")
        # biourja_mapping_loc=os.path.join(os.getcwd()+r"\Other Files\Template_File\Biourja_mapping1.xlsx")
        # mapping_loc=os.getcwd()+r"\Other Files\Templates\Biourja_mapping1.xlsx"
        # ---------------------------------------------------------------------------

        # We dont require input sheet2
        #----------------------------------------------Input Files and Templates
        # input_sheet2= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Input'+f'\\BS Bulk {month}{day}.xlsx'
        
        input_sheet= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Input'+f'\\AR Aging Bulk {month}{day}.xlsx'
        output_location = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Output'
        biourja_mapping_loc = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Template_File'+f'\\Biourja_mapping.xlsx'
        template_loc = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Template_File'+f'\\AR Aging Bulk Template.xlsx'
        grp_sheet_loc = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Template_File'+f'\\Group_mapping.xlsx'
        
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")  
        if not os.path.exists(biourja_mapping_loc):
            return(f"{biourja_mapping_loc} Excel file not present")  
        if not os.path.exists(template_loc):
            return(f"{template_loc} Excel file not present")    
        if not os.path.exists(grp_sheet_loc):
            return(f"{grp_sheet_loc} Excel file not present")

        # ----------------------------------------------------------------------------------------------------  

        wb = xlOpner(input_sheet)
        # wb = xw.Book("Test_file.xlsx")
        time.sleep(3)
        # Make Copy of Original Sheet and work on it
        original_tab= wb.sheets[0]
        original_tab.name = "Original_Data"
        original_tab.api.Copy(After=wb.api.Sheets(len(wb.sheets)))
        inp_sht = wb.sheets[len(wb.sheets)-1]
        inp_sht.name = "Custom_Sheet"
        inp_sht.activate()
        inp_sht.autofit()
        
        # Make Headers Bold
        header_range = inp_sht.range('1:1')
        header_range.api.Font.Bold = True
        
        #--------------------------------------------getting column indexes to delete------------------------------
        col_list=get_col_list(inp_sht)
        hash_col = col_list.index("#")
        hash_col_letter = num_to_col_letters(hash_col+1)
        inp_sht.range(f"{hash_col_letter}:{hash_col_letter}").delete()
        
        col_list=get_col_list(inp_sht)
        blanket_col = col_list.index("Blanket Agreement")
        blanket_col_letter = num_to_col_letters(blanket_col+1)
        inp_sht.range(f"{blanket_col_letter}:{blanket_col_letter}").delete()
        
        col_list=get_col_list(inp_sht)
        instal_col = col_list.index("Instal. No.")
        instal_col_letter = num_to_col_letters(instal_col+1)
        inp_sht.range(f"{instal_col_letter}:{instal_col_letter}").delete()
        
        col_list=get_col_list(inp_sht)
        payment_col = col_list.index("Payment Method Code")
        payment_col_letter = num_to_col_letters(payment_col+1)
        inp_sht.range(f"{payment_col_letter}:{payment_col_letter}").delete()
        
        col_list=get_col_list(inp_sht)
        dunning_col = col_list.index("Dunning Term Code")
        dunning_col_letter = num_to_col_letters(dunning_col+1)
        inp_sht.range(f"{dunning_col_letter}:{dunning_col_letter}").delete()
        
        col_list=get_col_list(inp_sht)
        doubt_col = col_list.index("Doubtful Debt")
        doubt_col_letter = num_to_col_letters(doubt_col+1)
        inp_sht.range(f"{doubt_col_letter}:{doubt_col_letter}").delete()
        
        col_list=get_col_list(inp_sht)
        bp_code_col = col_list.index("Consolidated BP Code")
        bp_code_col_letter = num_to_col_letters(bp_code_col+1)
        inp_sht.range(f"{bp_code_col_letter}:{bp_code_col_letter}").delete()
        
        col_list=get_col_list(inp_sht)
        bp_name_col = col_list.index("Consolidated BP Name")
        bp_name_col_letter = num_to_col_letters(bp_name_col+1)
        inp_sht.range(f"{bp_name_col_letter}:{bp_name_col_letter}").delete()
        
        
        # --------------------------------------------Fill up Blanks---------------------------------------------------------
        
        last_row_A = inp_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row
        last_row_D = inp_sht.range(f'D'+ str(inp_sht.cells.last_cell.row)).end('up').row
        
        last_row = last_row_A if last_row_A > last_row_D else last_row_D
        
        blank_cells= inp_sht.api.Range(f"A2:B{last_row}").SpecialCells(win32c.CellType.xlCellTypeBlanks).Select()
        wb.selection.api.FormulaR1C1 = "=+R[-1]C"
        
        # ----------------------------------------Copy 0-10 Column values-------------------------------------------------------
        col_list=get_col_list(inp_sht)
        future_remit_col = col_list.index("Future Remit")
        future_remit_col_letter = num_to_col_letters(future_remit_col+1)
        
        zero_ten_col = col_list.index("0 - 10")
        zero_ten_col_letter = num_to_col_letters(zero_ten_col+1)
        
        inp_sht.api.AutoFilterMode=False
        inp_sht.api.Range(f"A1:W{last_row}").AutoFilter(Field:=future_remit_col+1, Criteria1="=")
        inp_sht.api.Range(f"A1:W{last_row}").AutoFilter(Field:=zero_ten_col+1, Criteria1="<>")
        
        # inp_sht.api.Range(f"L1:L{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
                
        _,_,disc_ranges=row_range_calc(zero_ten_col_letter,inp_sht,wb)
        zero_ten_list=range_divider(disc_ranges,disc_ranges,zero_ten_col_letter,zero_ten_col_letter)
        remit_list=range_divider(disc_ranges,disc_ranges,future_remit_col_letter,future_remit_col_letter)
        
        for l_range,r_range in zip(zero_ten_list, remit_list):
            print(f"{l_range} and {r_range}")
            inp_sht.range(l_range).copy()
            inp_sht.range(r_range).paste()
        
        # delete 0-10 column now
        inp_sht.range(f"{zero_ten_col_letter}:{zero_ten_col_letter}").delete()

        inp_sht.api.AutoFilterMode=False
        
        #rename future remist column header
        inp_sht.range(f"{future_remit_col_letter}1").value = '0 - 10'
        
        
        ########################    Perform Shifting Here   ###############################
        
        if messagebox.showinfo("Prompt box",'Press ok after shifting is done'):
            print("promt clicked")
        else:
            return "Process aborted by user" 
        #######################################################
        
    
        # ------------------------------    Making Sub total using Macro    ----------------------------------
        
        # inp_sht.range('A1').api.Subtotal(GroupBy=2, Function=xlSum, TotalList=[2, 8, 9, 10, 11, 12, 13],
        #                        Replace=True, PageBreaks=True, SummaryBelowData=True)
        
        # In xlwings, when accessing the underlying Range object's API, you need to use the method directly on the .api attribute of the range object. The Function parameter for Subtotal corresponds to an integer value for the function being applied. In this case, 3 represents xlSum in the Excel object model.
        # inp_sht.range(f'A1:N815').api.Subtotal(GroupBy=2, Function=3, TotalList=[2, 8, 9, 10, 11, 12, 13],
        #                     Replace=True, PageBreaks=False, SummaryBelowData=True)
        
        # # Show Levels for RowGroups
        # inp_sht.api.Outline.ShowLevels(RowLevels=2)
        
        # --------------------------------- Making Seperate Df of Demrrauge Invoices -------------------------------------
        inp_sht.api.AutoFilterMode=False
        
        ## For test
        # inp_sht= wb.sheets["Data After Shifting"]
        
        sheet_data = inp_sht.used_range.value
    
        raw_df = pd.DataFrame(sheet_data[1:], columns=sheet_data[0])
        raw_df["Customer Name"]=raw_df["Customer Name"].str.upper()
        raw_df['Due Date'] = pd.to_datetime(raw_df['Due Date'])
        
        # Dropping Rows where Posting Date is None , this will ensure rows with Total will be dropped
        raw_df.dropna(subset=['Posting Date'], inplace=True)
        sorted_df = raw_df.sort_values(by='Due Date', ascending=False)
        
        removal_factors = ['DEMURRAGE', 'FACTORING', 'FREIGHT']
        temp_dataframe = pd.DataFrame()
        for index, row in sorted_df.iterrows():
            if(row['BP Ref. No.']!=None) and (row['BP Ref. No.']!='') :
                temp_str = row['BP Ref. No.'].strip().upper()
            else:
                continue
            
            # continue check if any removal factor is present in temp_str
            if any(factor.upper() in temp_str for factor in removal_factors):
                # Add the row to temp_dataframe
                temp_dataframe = temp_dataframe.append(row)
        # Temp Dataframe contains demmurage , frieght etc 
        print(temp_dataframe)
        
        # Removing Row with Total
        # rows_to_remove1 = []
        # totals_df=pd.DataFrame()
        # # Compare each row and append matching rows to temp_dataframe
        # for index, row in sorted_df.iterrows():
        #     customer_name = row['Customer Name'].strip().upper()
        #     if "TOTAL" in customer_name:
        #         totals_df = totals_df.append(row)
        #         rows_to_remove1.append(index)   # Store the index of the row to remove
        # sorted_df = sorted_df.drop(rows_to_remove1)
        
        # Remove the matching rows from grp_df
        # grp_df = temp_dataframe.drop(rows_to_remove1)
        
        # =======================================Making sub total of demurrage ==============================================
        non_products_df = temp_dataframe.groupby('Customer Name', sort=False)[['Balance Due', '0 - 10', '11 - 30', '31 - 60', '61 - 90', '91+']].sum().reset_index()
        
        non_products_df = non_products_df.sort_values(by='Customer Name', ascending=True)
        
        non_products_df[['Balance Due', '0 - 10', '11 - 30', '31 - 60', '61 - 90', '91+']] = non_products_df[['Balance Due', '0 - 10', '11 - 30', '31 - 60', '61 - 90', '91+']].round(2)
        
        # Make a new sheet rename it and paste
        summary_sheet=wb.sheets.add("Demurrage,Freight...",after=wb.sheets[0])
        summary_sheet.activate()
        
        #================ Make new sheet for Demurrage and Non products===============================
        summary_sheet.range(f"A1").options(index = False).value=non_products_df
        header_range = summary_sheet.range('1:1')
        header_range.api.Font.Bold = True
        summary_sheet.autofit()
        # ==================================== Making Subtotal table of All AR Companies ====================
        
        # Making Subtotal of of all values
        grp_df = sorted_df.groupby('Customer Name', sort=False)[['Balance Due', '0 - 10', '11 - 30', '31 - 60', '61 - 90', '91+']].sum().reset_index()
        
        grp_df = grp_df.sort_values(by='Customer Name', ascending=True)
        
        grp_df[['Balance Due', '0 - 10', '11 - 30', '31 - 60', '61 - 90', '91+']] = grp_df[['Balance Due', '0 - 10', '11 - 30', '31 - 60', '61 - 90', '91+']].round(2)
        
        # ================================================================================================
        
        # Get list of cis concern companies
        # mapping_loc=os.getcwd()+r"\Other Files\Templates\Biourja_mapping1.xlsx"
        
        mapping_df=pd.read_excel(biourja_mapping_loc)
        time.sleep(3)
        # Seperating cis concern company from original company
        cis_concern_df = pd.DataFrame()      
        rows_to_remove = []

        # Compare each row and append matching rows to temp_dataframe
        for index, row in grp_df.iterrows():
            customer_name = row['Customer Name'].strip().upper()
            matching_row = mapping_df[mapping_df['CompanyName'].str.strip().str.upper() == customer_name]
            if not matching_row.empty:
                cis_concern_df = cis_concern_df.append(row)
                # rows_to_remove.append(index)  # Store the index of the row to remove

        # We are not removing rows as they will be handled later
        # Remove the matching rows from grp_df
        # grp_df = grp_df.drop(rows_to_remove)
        
        # Make a new sheet rename it and paste
        summary_sheet=wb.sheets.add("Summary_Sheet",after=wb.sheets[0])
        summary_sheet.activate()
        
        # Paste All AR product+ Cis concern companies
        summary_sheet.range(f"A1").options(index = False).value=grp_df
        header_range = summary_sheet.range('1:1')
        header_range.api.Font.Bold = True
        summary_sheet.autofit()
        
        # Make a new sheet for cis concern companies
        summary_sheet=wb.sheets.add("Cis_Concern_Companies",after=wb.sheets[0])
        summary_sheet.activate()
        summary_sheet.range(f"A1").options(index = False).value=cis_concern_df
        header_range = summary_sheet.range('1:1')
        header_range.api.Font.Bold = True
        summary_sheet.autofit()
        
        # =============================== Adding Template and filling data 
        
        # template_loc=os.path.join(os.getcwd(),r"\Other Files\Templates\AR Aging Bulk Template.xlsx")
        
        wb_temp = xlOpner(template_loc)
        # wb_temp.write= xlOpner("AR Aging Bulk Template.xlsx")
        
        # ---------------------------------------------------------------------------------
        bulk_tab= wb_temp.sheets["Bulk(2)"]
        bulk_tab.api.Copy(After=wb.api.Sheets(len(wb.sheets)))
        bulk_tab_it = wb.sheets[len(wb.sheets)-1]
        bulk_tab_it.name = "Bulk_Data(IT)"
        
        #----------------------------------------- Change Date ----------------------------------
        intial_date = bulk_tab_it.range("B3").value.split("To")[0].strip()
        last_date = bulk_tab_it.range("B3").value.split("To")[1].strip()
        intial_date_xl = f"01-01-{year}"
        last_date = f"{month}-{day}-{year}"
        xl_input_Date = intial_date_xl + f" To " + last_date
        bulk_tab_it.range("B3").value = xl_input_Date
    
        
        #--------------------------Deleting Existing Data from template till INELIGIBALE ACCOUNTS
        bulk_tab_it.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = bulk_tab_it.api.Application.ActiveCell.Address.replace("$","")# Get the index of row with INELIGIBLE ACCOUNTS
        
        #Taking integer part of index found
        brow_value = re.findall("\d+",bcell_value)[0]
        
        #Delete part below INELIGIBLE ACCOUNTS RECEIVABLE
        bulk_tab_it.range(f"B{int(brow_value)+1}").expand('table').api.Delete()
        
        #Delete Part above INELIGIBLE ACCOUNTS RECEIVABLE
        bulk_tab_it.range(f"B9:J{int(brow_value)-1}").api.Delete()
        
        
        #-------------------------------------- Delete Data in Excluded AR Table--------------------------
        delete_row_end = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        delete_row_end2 = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
        
        bulk_tab_it.range(f"{delete_row_end2}:{delete_row_end2}").insert()# Insert one row before deleting so the lines won't collapse
        bulk_tab_it.range(f"{delete_row_end2+1}:{delete_row_end+1}").api.Delete()


        # bulk_tab_it.api.Range(f"B8:C{ini-1}").Copy(bulk_tab_it.range(f"B100").api)
        
        summary_tab= wb.sheets["Summary_Sheet"]
        last_row_summary_A = summary_tab.range(f'A'+ str(summary_tab.cells.last_cell.row)).end('up').row
        
        # Pasting Data From summary sheet to template at row 100 on order to handle insert row  case
        summary_tab.api.Range(f"A2:B{last_row_summary_A}").Copy(bulk_tab_it.range(f"B100").api)
        
        bulk_tab_it.activate()
        bulk_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
        bulk_tab_it.range(f"B9").api.EntireRow.Select()
        
        # Insert Copied B100 Data with Insert Row + Data
        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
        
        # Getting 1st row of data inserted at after the excluded AR part to delete it
        ini2 = bulk_tab_it.range(f'B'+ str(summary_tab.cells.last_cell.row)).end('up').end('up').row
        
        # Insert rows
        bulk_tab_it.range(f"B{ini2}").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        
        # Getting last row of B now = cell with Ineligible account receivable
        ini2 = bulk_tab_it.range(f'B'+ str(summary_tab.cells.last_cell.row)).end('up').end('up').row
        
        # Filldown First row Formulas
        bulk_tab_it.api.Range(f"D8:J{ini2-1}").Select()
        wb.app.api.Selection.FillDown()
        
        # Deleting 1st row from old Template since we have used the formula from it , using win32 to delete row with content
        bulk_tab_it.api.Range(f"8:8").EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bulk_tab_it.api.Range(f"B{ini2-1}").Font.Bold = True
        
        # Next Copy Part
        summary_tab.api.Range(f"C2:G{last_row_summary_A}").Copy(bulk_tab_it.range(f"E8").api)
        
        # Multiply with -1  No need to multiply with -1 already handled in SAP
        # bulk_tab_it.api.Range(f"J1").Copy()
        # bulk_tab_it.api.Range(f"C8:C{ini2-2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)
        # bulk_tab_it.api.Range(f"E8:I{ini2-2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlMultiply)

        # ==============================================================================================
        # We Apply Vlookup here to remove duplicate companies if any , will apply final vlookup later
        # grp_sheet_loc=os.path.join(os.getcwd()+r"\Other Files\Template_File\Group_mapping.xlsx")
        
        grp_wb=xlOpner(grp_sheet_loc)
    
        time.sleep(3)
        bulk_tab_it.activate()
        bulk_tab_it.api.Range(f"L8").Value="=+XLOOKUP(B8,'[Group_mapping.xlsx]Sheet1'!$A:$A,'[Group_mapping.xlsx]Sheet1'!$B:$B,0)"

        bulk_tab_it.api.Range(f"L8:L{ini2-2}").Select()
        
        # Fill Down Xlookup applied at L8
        wb.app.api.Selection.FillDown()
        bulk_tab_it.api.Range(f"L7").Select()
        
        # Give Name to column Since Filtering Won't be allowed else
        bulk_tab_it.api.Range(f"L6").Value = "Xlookup"
        bulk_tab_it.api.AutoFilterMode=False
        bulk_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:='=0')
        
        # Getting last row To Identify any xlookup with zero(False)
        lst_row_with_zero = bulk_tab_it.range(f'L'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        if lst_row_with_zero != 8:
            sp_address= bulk_tab_it.api.Range(f"L8:L{lst_row_with_zero}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
            
            # This will return index of first row with 0
            initial_rw_with_zero = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        else:
            # Xlookup with zero not present
            initial_rw_with_zero = 8
        
        # Remove All Zeroes
        bulk_tab_it.range(f"L{initial_rw_with_zero}").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
        try:
            # Getting All rows back
            bulk_tab_it.api.Range(f"L7").AutoFilter(Field:=1)
        except:
            pass
        
        # Finding if any company listed twice
        font_colour,Interior_colour = conditional_formatting(range=f"L:L",working_sheet=bulk_tab_it,working_workbook=wb)
        bulk_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        sp_lst_row = bulk_tab_it.range(f'L'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        sp_address= bulk_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        if sp_lst_row ==int(sp_initial_rw):
            print("no data to filter")
            grp_cm_list2=[]
        else:
            bulk_tab_it.range(f"L{sp_initial_rw}:L{sp_lst_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
            bulk_tab_it.api.Range(f"B100")._PasteSpecial(Paste=-4163)
            grp_cm_list = bulk_tab_it.range(f"B100").expand('down').value
            bulk_tab_it.range(f"B100").expand('down').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            grp_cm_list2 = list(set(grp_cm_list))
            bulk_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        
        # Getting Row number where we will update formula
        val_row = bulk_tab_it.range(f'C'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        if len(grp_cm_list2)>0:
            for i in range(len(grp_cm_list2)):
                bulk_tab_it.api.Range(f"L7").Select()
                bulk_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:=[grp_cm_list2[i]])
                sp_lst_row = bulk_tab_it.range(f'L'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
                sp_address= bulk_tab_it.api.Range(f"L8:L{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
                if bulk_tab_it.range(f"C{sp_initial_rw}").value + bulk_tab_it.range(f"C{sp_lst_row}").value<0:
                    bulk_tab_it.range(f"{sp_initial_rw}:{sp_lst_row}").api.EntireRow.Copy()
                    bulk_tab_it.range(f"{val_row+3}:{val_row+3}").api.EntireRow.Select()
                    wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                    bulk_tab_it.range(f"{sp_initial_rw}:{sp_lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                else:
                    print("second case")

            bulk_tab_it.api.Cells.FormatConditions.Delete()
            bulk_tab_it.api.AutoFilterMode=False
            
        # val_row = bulk_tab_it.range(f'C'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        # Comment deletion part if required
        bulk_tab_it.api.Range(f"L:L").EntireColumn.Delete()
        
        #====================================================Finding Negative AR==================================================
        bulk_tab_it.api.AutoFilterMode=False
        font_colour,Interior_colour = conditional_formatting2(range=f"C8:C{ini2-2}",working_sheet=bulk_tab_it,working_workbook=wb)
        bulk_tab_it.api.Range(f"C7").AutoFilter(Field:=2, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        
        # Getting last row with negative value , then finding splitted ranges , then finding the first row among them
        sp_lst_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
        sp_address= bulk_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        negative_check=True
        if int(sp_initial_rw)==6:
            negative_check=False
            pass
        elif int(sp_lst_row) ==int(sp_initial_rw):
            bulk_tab_it.range(f"B{sp_initial_rw}").expand("right").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(bulk_tab_it.range(f"B100").api)
        else:    
            bulk_tab_it.range(f"B{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(bulk_tab_it.range(f"B100").api)

        if(negative_check):
        # Copy to b100 then paste to A{val_row+3} = > excluded AR part
            bulk_tab_it.range(f"B100").expand('down').api.EntireRow.Copy()
            bulk_tab_it.range(f"A{val_row+3}").api.EntireRow.Select()#Formula row +3
            wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
            wb.app.api.CutCopyMode=False
        
        # Since we inserted new rows , B100 rows will be shifted finding first row index to delete the data
        rw_utility=bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
        if rw_utility==6:
            pass
        elif val_row+3 ==rw_utility:
            # Only one row inserted
            rw_utility=bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
            bulk_tab_it.range(f"B{rw_utility}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        else:
            # More than one row inserted
            bulk_tab_it.range(f"B{rw_utility}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        # Now since copy done , Deleting from the main AR
        if int(sp_initial_rw)==6:
            # No negative data found , no need to delete
            pass
        elif int(sp_lst_row) ==int(sp_initial_rw):
            # Deleting this rows from AR ; first row == last row => Only one negative value. 
            bulk_tab_it.range(f"B{sp_initial_rw}").expand('right').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        else:    
            bulk_tab_it.range(f"B{sp_initial_rw}").expand('table').api.EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        bulk_tab_it.api.AutoFilterMode=False
        
        # Cis Concern Possiblities
        # 1. No Cis Concern present   +No Negative AR
        # 2. No Cis Concern present   +Negative Ar Present
        # 3. Only 1 Cis Concern 
        # 4. Only 1 Cis Concern      +Negative AR
        # 5. More than 1 Cis Concern
        # 6. More than 1 Cis Concern +Negative AR
        
        #========================Shifting Cis Concern Companies to excluded AR=============================
        if len(cis_concern_df)>0:
            # biourja_mapping_loc=os.path.join(os.getcwd()+r"\Other Files\Template_File\Biourja_mapping1.xlsx")
            company_wb=xlOpner(biourja_mapping_loc)
 
            time.sleep(3)
            company_sheet = company_wb.sheets[0] 
            company_names = company_sheet.range(f"A2").expand('down').value

            company_names = [names.strip() for names in company_names]
            
            # Copy Company names and paste at B100 of BulkItTab
            company_sheet.range(f"A2").expand('down').api.Copy(bulk_tab_it.range(f"B100").api)
            bulk_tab_it.api.Cells.FormatConditions.Delete()
            bulk_tab_it.activate()
            
            # Finding Duplicate between Cis concern name pasted at B100 and Present in AR report.
            font_colour,Interior_colour = conditional_formatting(range=f"B:B",working_sheet=bulk_tab_it,working_workbook=wb)
            bulk_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

            sp_lst_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
            
            # Return the address of visible cells ex : 6:6 5:5 10:10 , may or may not be continous
            sp_address= bulk_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            sp_initial_rw1 = sp_address.split(",")[0].split(":")[1].replace("$","")

            # No Cis Concern present
            if(sp_initial_rw==6 or bulk_tab_it.range(f"B{int(sp_initial_rw)}").value==None or bulk_tab_it.range(f"B{int(sp_initial_rw)}").value==''):
                print("No Cis Concern Companies")
                
                added_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
                bulk_tab_it.range(f"b{added_row}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                
            # Only 1 Cis Concern Company Present
            elif(sp_initial_rw==sp_initial_rw1):
                sp_initial_rw2=re.findall("\d+",sp_address.split(",")[1].replace("$","").split(":")[0])[0]

                if(bulk_tab_it.range(f"B{int(sp_initial_rw2)}").value==None or bulk_tab_it.range(f"B{int(sp_initial_rw2)}").value==''):
                    print("Only one cis concern row")
                    utility_sh=wb.sheets.add("utility_sheet",after=wb.sheets[-1])
                    bulk_tab_it.range(f"B{sp_initial_rw}").expand('right').api.Copy(utility_sh.range(f"B1").api)
                    utility_sh.range(f"B1").expand('table').api.EntireRow.Copy()
                    row_to_insert=bulk_tab_it.range(f'C'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
                    
                    bulk_tab_it.activate()
                    
                    value_row2 = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').end('up').row
                    
                    if bulk_tab_it.range(f"B{value_row2}").value=='Total':
                        bulk_tab_it.range(f"A{row_to_insert+3}").api.EntireRow.Select()
                    else:
                        bulk_tab_it.range(f"A{row_to_insert+1}").api.EntireRow.Select()
                        
                    wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                    
                    bulk_tab_it.range(f"B{sp_initial_rw}").expand('right').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                    utility_sh.delete()
                    bulk_tab_it.api.AutoFilterMode=False
                    bulk_tab_it.api.Cells.FormatConditions.Delete()
                    
                    added_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
                    bulk_tab_it.range(f"b{added_row}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                    added_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
                    bulk_tab_it.range(f"b{added_row}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            else:       
                # More Than One Cis Concern Company Present
                value_row2 = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').end('up').row

                bulk_tab_it.range(f"B{sp_initial_rw}").expand('table').api.Copy(bulk_tab_it.range(f"B150").api)

                bulk_tab_it.range(f"B150").expand('table').api.EntireRow.Copy()
                
                if bulk_tab_it.range(f"B{value_row2}").value=='Total':
                    value_row2 = bulk_tab_it.range(f'C'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').end('up').row+2
                bulk_tab_it.range(f"A{value_row2+1}").api.EntireRow.Select()
                wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                bulk_tab_it.range(f"B{sp_initial_rw}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)

                bulk_tab_it.api.AutoFilterMode=False
                bulk_tab_it.api.Cells.FormatConditions.Delete()
                
                added_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
                bulk_tab_it.range(f"b{added_row}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                added_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').row
                bulk_tab_it.range(f"b{added_row}").expand('table').api.SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Delete(win32c.DeleteShiftDirection.xlShiftUp)

        #  Cis Concern Completed Till here
        
        # ================================== Managing Demurrage and Non Products==================
        if non_products_df.size>0:
            non_products_df.reset_index(drop=True, inplace=True)
            non_products_df.insert(2,"> 10",non_products_df[['11 - 30','31 - 60','61 - 90','91+']].sum(axis=1))
            for column in non_products_df.columns:
                if non_products_df[column].dtype == 'float64':
                    non_products_df[column] *= -1
            non_products_df['As Per BS'] = non_products_df['Balance Due'] - non_products_df['0 - 10'] - non_products_df['> 10']


        bulk_tab_it.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = bulk_tab_it.api.Application.ActiveCell.Address.replace("$","")
        brow_value = re.findall("\d+",bcell_value)[0]

        if non_products_df.size>0:
            
            # Insert Blank Rows
            bulk_tab_it.api.Range(f"B{int(brow_value)+1}:B{int(brow_value)+len(non_products_df)}").EntireRow.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
            
            # Inserting Demurrage freight
            bulk_tab_it.range(f'B{int(brow_value)+1}').options(index = False,header=False).value = non_products_df
            bulk_tab_it.range(f'B{int(brow_value)+1}').expand('down').font.bold= False
            
            # Sorting
            bulk_tab_it.range(f"B8:J{int(brow_value)-1}").api.Sort(Key1=bulk_tab_it.range(f"B8:B{int(brow_value)-1}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
            bulk_tab_it.range(f'B{int(brow_value)+1}').expand('table').api.Sort(Key1=bulk_tab_it.range(f'B{int(brow_value)+1}').expand('down').api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

        if non_products_df.size>0:    
            for i in range(len(non_products_df['Customer Name'])):
                conditional_formatting(range=bulk_tab_it.range(f'B8').expand('table').get_address(),working_sheet=bulk_tab_it,working_workbook=wb)
                bulk_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
                bulk_tab_it.api.Range(f"B7").AutoFilter(Field:=1, Criteria1:=[non_products_df['Customer Name'][i]])
                sp_lst_row = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
                sp_address= bulk_tab_it.api.Range(f"B8:B{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]  
                int_check = bulk_tab_it.range(f"B{sp_initial_rw}").expand("table").get_address().split(":")[-1]
                lst_row = re.findall("\d+",int_check .replace("$","").split(":")[0])[0]
                
                # Shifting case
                if bulk_tab_it.range(f"C{sp_initial_rw}").value + bulk_tab_it.range(f"C{lst_row}").value<=1:
                    bulk_tab_it.range(f"{lst_row}:{lst_row}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                    in_rw = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
                    bulk_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.EntireRow.Copy()
                    bulk_tab_it.range(f"{in_rw+1}:{in_rw+1}").api.EntireRow.Select()
                    wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                    bulk_tab_it.range(f"{sp_initial_rw}:{sp_initial_rw}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                    bulk_tab_it.api.AutoFilterMode=False
                    bulk_tab_it.api.Cells.FormatConditions.Delete()
                else:
                    print("second case")
                    bulk_tab_it.api.AutoFilterMode=False
                    bulk_tab_it.api.Cells.FormatConditions.Delete()

        # Ineligible accounts check
        bulk_tab_it.api.Cells.Find(What:="INELIGIBLE ACCOUNTS RECEIVABLE", After:=bulk_tab_it.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = bulk_tab_it.api.Application.ActiveCell.Address.replace("$","")
        brow_value = re.findall("\d+",bcell_value)[0]
    
        if bulk_tab_it.range(f"B{int(brow_value)+1}").value!=None:# Check if Any Ineligible account remains , if not delete ineligible account heading
            pass
        else:
            bulk_tab_it.range(f"{brow_value}:{brow_value}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)

        #==============================================Updating formula================================
        
        formula_row = bulk_tab_it.range(f'C'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').end('up').row
        pre_row = bulk_tab_it.range(f"C{formula_row}").end('up').row
        fst_rng = bulk_tab_it.range(f"C8").expand("down").get_address().replace("$","")
        mid_range = bulk_tab_it.range(f"C{formula_row}").formula.split("+")[-1].split("-")[0]
        bulk_tab_it.range(f"C{formula_row}").formula = f"=+SUM({fst_rng})+{mid_range}-C{pre_row}"
        
        #=====================================Formatting Table font size etc===========================
        
        bulk_tab_it.range(f"J8").expand('down').number_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        bulk_tab_it.range(f"J8").expand('down').font.size = 9
        
        extreme_last_row = bulk_tab_it.range(f'C'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        bulk_tab_it.range(f"B8:L{extreme_last_row}").font.size = 9
        
        #=============================================Balance Sheet part==================================
        
        # input_sheet2= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Input'+f'\\BS Bulk {month}{day}.xlsx'
        # retry=0
        # while retry < 10:
        #     try:
        #         bulk_wb = xw.Book(input_sheet2,update_links=False) 
        #         break
        #     except Exception as e:
        #         time.sleep(5)
        #         retry+=1
        #         if retry ==10:
        #             raise e 

        # bs_tab = bulk_wb.sheets[0]   
        # bs_tab.activate()
        # bs_tab.range(f"A1").select()     
        # bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        # cell_value = bs_tab.api.Application.ActiveCell.Address.replace("$","")
        # row_value = re.findall("\d+",cell_value)[0] 
        # bs_tab.api.Cells.Find(What:="accounts receivable", After:=bs_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        # cell_value2 = bs_tab.api.Application.ActiveCell.Address.replace("$","")
        # row_value2 = re.findall("\d+",cell_value2)[0]
        # bs_tab.api.Range(f"B{row_value}:C{int(row_value2)-1}").Copy(bs_tab.range(f"I1").api)

        # bs_tab.api.Range(f"J1").AutoFilter(Field:=2, Criteria1:=['=0.00'],Operator:=2,Criteria2:="=0.01")
        # sp_lst_row = bs_tab.range(f'I'+ str(bs_tab.cells.last_cell.row)).end('up').row
        # sp_address= bs_tab.api.Range(f"I2:I{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        # bs_tab.range(f"I{sp_initial_rw}").expand("table").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
        # bs_tab.api.AutoFilterMode=False 
        # time.sleep(1)
        # bs_total = round(sum(bs_tab.range(f"J2").expand('down').value),2)
        # bs_tab.range(f"I2").expand("table").copy(bulk_tab_it.range(f"L8"))
        
        # ===========================================================================================================
        
        #============================================ Updaing Xlook ==================================================
        # grp_sheet_loc=os.path.join(os.getcwd()+r"\Other Files\Template_File\Group_mapping.xlsx")
        
        bulk_tab_it.activate()
        bulk_tab_it.api.Range(f"L8").Value="=+XLOOKUP(B8,'[Group_mapping.xlsx]Sheet1'!$A:$A,'[Group_mapping.xlsx]Sheet1'!$B:$B,0)"
        last_row_for_vlookup = bulk_tab_it.range(f'B'+ str(bulk_tab_it.cells.last_cell.row)).end('up').end('up').end('up').end('up').row
        bulk_tab_it.api.Range(f"L8:L{last_row_for_vlookup}").Select()
        wb.app.api.Selection.FillDown()
        bulk_tab_it.api.Range(f"L7").Select()
        bulk_tab_it.api.Range(f"L6").Value = "Xlookup"
        bulk_tab_it.range('6:6').api.Font.Bold = True
        bulk_tab_it.api.AutoFilterMode=False

        bulk_tab_it.api.Range(f"L7").AutoFilter(Field:=1, Criteria1:='=0')
        
        # Getting last row To Identify any xlookup with zero(False)
        lst_row_with_zero = bulk_tab_it.range(f'L'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        if lst_row_with_zero != 8:
            sp_address= bulk_tab_it.api.Range(f"L8:L{lst_row_with_zero}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
            
            # This will return index of first row with 0
            initial_rw_with_zero = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        else:
            # Xlookup with zero not present
            initial_rw_with_zero = 8
        
        # Remove All Zeroes
        bulk_tab_it.range(f"L{initial_rw_with_zero}").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()
        try:
            # Getting All rows back
            bulk_tab_it.api.Range(f"L7").AutoFilter(Field:=1)
        except:
            pass
        bulk_tab_it.api.AutoFilterMode=False
        
        # ================================ Adding Number Format ====================================
        
        extreme_last_row = bulk_tab_it.range(f'C'+ str(bulk_tab_it.cells.last_cell.row)).end('up').row
        extreme_first_row = bulk_tab_it.range(f'C1').end('down').end('down').row
        
        bulk_tab_it.range(f"C{extreme_first_row}:J{extreme_last_row}").number_format='0.00'
        # =============================== Deleting Extra Tabs =============================

        try:
            wb.sheets["Summary_Sheet"].delete()
        except:
            pass
        
        #------------------------------ Colouring Tabs ----------------------------------------
        original_tab= wb.sheets["Original_Data"]
        data_tab= wb.sheets["Custom_Sheet"]
        
        tablist={data_tab:win32c.ThemeColor.xlThemeColorAccent5,bulk_tab_it:win32c.ThemeColor.xlThemeColorAccent4,original_tab:win32c.ThemeColor.xlThemeColorAccent6}
        for tab,color in tablist.items():
                tab.activate()
                tab.api.Tab.ThemeColor = color
                tab.autofit()
                tab.range(f"A1").select()

        #-------------------------------- Shift Tabs ----------------------------------------------
        before_sheet = wb.sheets[0]
        # Move "Data_Sheet" before the sheet at index 1
        original_tab.api.Move(Before=before_sheet.api)
        
        before_sheet = wb.sheets[1]
        # Move "Data_Sheet" before the sheet at index 1
        data_tab.api.Move(Before=before_sheet.api)
        
        before_sheet = wb.sheets[2]
        # Move "Data_Sheet" before the sheet at index 1
        bulk_tab_it.api.Move(Before=before_sheet.api)
        
        wb.save(f"{output_location}\\AR Aging Bulk {month}{day}-updated"+'.xlsx')
        wb.app.quit()
        return "Report Generated Successfully please check in output folder"
    except Exception as e:
        wb.app.quit()
        print("Exception caught during execution: ",e)
        logging.exception(f'Exception caught during execution: {e}')
        raise e


if __name__ == "__main__":
    try:
        # logfile = os.getcwd() + '\\' + 'logs' + '\\' + 'AR_AGEING_BBR.txt'
        # for handler in logging.root.handlers[:]:
        #     logging.root.removeHandler(handler)
        # logging.basicConfig(
        #     level=logging.INFO, 
        #     format='%(asctime)s [%(levelname)s] - %(message)s',
        #     filename=logfile)
        
        #######################Uncommment for Testing###################
        database="BUITDB_DEV"
        warehouse="BUIT_WH"
        receiver_email="amanullah.khan@biourja.com"
        job_name="Test AR AGEING BBR"
        job_id=np.random.randint(1000000,9999999)
        ################################################################
        
        today_date = date.today() - timedelta(days=0)
        
        # ----------------------------- For Debug Only -----------------------------
        input_date="01.15.2023"
        output_date="01.15.2023"
        
        ar_ageing_BBR(input_date,output_date)
    except Exception as e:
        print(e)