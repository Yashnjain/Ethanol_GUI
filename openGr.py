import xlwings as xw
import xlwings.constants as win32c
import os, time
from datetime import datetime, timedelta
import re
import pandas as pd

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

def conditional_formatting(columnvalue:str,working_sheet,working_workbook):

    try:

        font_colour = -16383844

        Interior_colour = 13551615

        working_sheet.api.Range(f"{columnvalue}:{columnvalue}").Select()

        working_workbook.app.selection.api.FormatConditions.AddUniqueValues()

        working_workbook.app.selection.api.FormatConditions(working_workbook.app.selection.api.FormatConditions.Count).SetFirstPriority()

 

        working_workbook.app.selection.api.FormatConditions(1).DupeUnique = win32c.DupeUnique.xlDuplicate

 

        working_workbook.app.selection.api.FormatConditions(1).Font.Color = font_colour

        working_workbook.app.selection.api.FormatConditions(1).Interior.Color = Interior_colour

        working_workbook.app.selection.api.FormatConditions(1).Interior.PatternColorIndex = win32c.Constants.xlAutomatic

        return font_colour,Interior_colour

    except Exception as e:

        raise e



def knockOffAmtDiff(curr,final, wb, input_sht, input_sht2, credit_col_letter, debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter, eth_trueup_col_letter=None):
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
            input_sht.range(f"{curr}:{final}").copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))

            input_sht.range(f"{curr}:{final}").delete()

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
    elif (abs(input_sht.range(f"{credit_col_letter}{curr}").value) - abs(input_sht2.range(f"{debit_col_letter}{final}").value))<10:
        #amt diff
        print(f"Moving {curr} to amount diff")
        amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row

        if input_sht==input_sht2:
            input_sht.range(f"{curr}:{final}").api.Copy()
            wb.activate()
            amt_diff_sht.activate()
            amt_diff_sht.range(f"A{amt_diff_last_row+1}").api.Select()
            wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
            amt_diff_sht.autofit()
            # input_sht.range(f"{i}:{i+1}").copy(amt_diff_sht.range(f"A{amt_diff_last_row+1}"))

            input_sht.range(f"{curr}:{final}").delete()
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
            input_sht2.range(f"{final}:{final}").delete()

        curr-=1
    else:
        #line for ethnaol accrual tab
        print(f'current line {curr} remains here for ethanol accrual tab having mrn no.{input_sht.range(f"{mrn_col_letter}{curr}")}')
    return curr

def row_range_calc(input_sht, wb):
    sp_lst_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row

    sp_address= input_sht.api.Range(f"A2:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address

    sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

    row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

    while row_range[-1]!=sp_lst_row:

        sp_lst_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row

        sp_address.extend(input_sht.api.Range(f"A{row_range[-1]}:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)

        # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

        # row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
        
    
    sp_address = sp_address.replace("$","").split(",")
    init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
    sublist = []
    flat_list = [item for sublist in init_list for item in sublist]
    return flat_list, sp_lst_row,sp_address

def openGr2(input_date, output_date):
    try:
        input_datetime = datetime.strptime(input_date, "%m.%d.%Y")
        month = input_datetime.month
        day = input_datetime.day
        j_loc = r"J:\India\BBR\IT_BBR\Reports"
        # curr_loc = os.getcwd()
        # input_sheet= curr_loc+r'\Raw Files'+f'\\Open GR {month}{day}.xlsx'
        input_sheet= j_loc+r'\Open GR\Raw Files'+f'\\Open GR {month}{day}.xlsx'
        output_location = j_loc+r'\Open GR\Output Files' 
        # output_location = curr_loc+r'\Output Files' 
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
        #make copy of Sheet1
        wb.sheets["Input"].copy(name="Input_Main", after=wb.sheets["Input"])
        input_sht = wb.sheets["Input_Main"]
        
       
        #Deleting extras
        input_sht.range("A:A").delete()
        input_sht.range(f'1:{input_sht.range("A1").end("down").end("down").row-1}').delete()

        #Checking Opening Balance
        curr_col_list = input_sht.range("A1").expand('right').value
        balance_row = input_sht.range(f'{num_to_col_letters(len(curr_col_list))}1').end('down').row -1
        balance = input_sht.range(f"{num_to_col_letters(curr_col_list.index('Balance')+1)}{balance_row}").value

        reco_sht = wb.sheets["Reco"]
        reco_last_row = reco_sht.range(f'A'+ str(reco_sht.cells.last_cell.row)).end('up').row
        reco_col_list = reco_sht.range("A1").expand('right').value
        reco_a_list = reco_sht.range(f"A1:A{reco_last_row}").value
        input_total = wb.sheets["Input"].range(f"AB{wb.sheets['Input'].range(f'AC'+ str(wb.sheets['Input'].cells.last_cell.row)).end('up').row}").address
        #Updating reco input sheet value
        reco_sht.range("B8").formula = f"=+'Input'!{input_total}"
        # if balance != reco_sht.range(f'{num_to_col_letters(len(reco_col_list))}{reco_a_list.index("Open MRN as Per BS")+1}').value:
        #     return "Opening blanace of Input Sheet not balanced with Reco sheet"


        







        #Extra Column deletion logic
        req_col_list = ["Date", "Cost Center", "Terminal", "Voucher No", "Name", "Vendor Ref", "Pur VNo", "MRN No:", "BOL Number", "Rail Car/Truck #",
        "Narration"	"Remarks", "Debit Amount", "Credit Amount", "Balance"]
        
        i=0
        while len(req_col_list) <=len(curr_col_list):
            curr_col = num_to_col_letters(i+1)
            if input_sht.range(f"{curr_col}1").value not in req_col_list:
                input_sht.range(f"{curr_col}:{curr_col}").delete()
                i-=1
            curr_col_list = input_sht.range("A1").expand('right').value
            i+=1
        #Delete extra rows of total in starting
        to_be_deleted = input_sht.range("A1").end('down').row
        input_sht.range(f"2:{to_be_deleted-1}").delete()
        #Sorting by railcar
        curr_last_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
        curr_last_col = len(curr_col_list)
        curr_last_col_letter = num_to_col_letters(curr_last_col)
        railcar_col = curr_col_list.index("Rail Car/Truck #")
        railcar_col_letter = num_to_col_letters(railcar_col+1)
        

        input_sht.range(f"A1:{curr_last_col_letter}{curr_last_row}").api.Sort(Key1=input_sht.range(f"{railcar_col_letter}1:{railcar_col_letter}{curr_last_row}").api,
            Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

        # #Removing Extra total
        to_be_deleted_final = input_sht.range(f'B'+ str(input_sht.cells.last_cell.row)).end('up').row
        to_be_deleted_init = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row

        input_sht.range(f"{to_be_deleted_init+1}:{to_be_deleted_final+5}").delete()#+% for deleting extra line border
        input_sht.copy(name = "Input_Main2", after = wb.sheets["Input_Main"])
        input_sht = wb.sheets["Input_Main2"]

        voucher_col = curr_col_list.index("Voucher No")
        voucher_col_col_letter = num_to_col_letters(voucher_col+1)
        
        mrn_col = curr_col_list.index("MRN No:")
        mrn_col_letter = num_to_col_letters(mrn_col+1)

        date_col = curr_col_list.index("Date")
        date_col_letter = num_to_col_letters(date_col+1)


        debit_col = curr_col_list.index("Debit Amount")
        debit_col_letter = num_to_col_letters(debit_col+1)

        credit_col = curr_col_list.index("Credit Amount")
        credit_col_letter = num_to_col_letters(credit_col+1)

        bol_col = curr_col_list.index("BOL Number")
        bol_col_letter = num_to_col_letters(bol_col+1)


        last_row = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
        curr_month_num =datetime.strptime(input_date,"%m.%d.%Y").month
        curr_month = datetime.strftime(datetime.strptime(input_date,"%m.%d.%Y"), "%b")
        prev_month = datetime.strftime((datetime.strptime(input_date,"%m.%d.%Y").replace(day=1) -timedelta(days=1)), "%b")


        #Adding all sheets at once #Logic avoided as these sheets presaent from previous file
        # knock_off_sht = wb.sheets.add("Knocked Off",after=wb.sheets[-1])
        # amt_diff_sht = wb.sheets.add("Amount Diff",after=wb.sheets[-1])
        # diff_month_sht = wb.sheets.add(f"{prev_month} MRN booked in {curr_month}",after=wb.sheets[-1])

        knock_off_sht = wb.sheets("Knocked Off")
        amt_diff_sht = wb.sheets("Amount Diff")
        try:
            diff_month_sht = wb.sheets(f"{prev_month} MRN booked in {curr_month}")
        except:
            diff_month_sht = wb.sheets.add(f"{prev_month} MRN booked in {curr_month}",after=amt_diff_sht)

        #Adding headers in all new sheets
        input_sht.range(f"A1").expand("right").copy(knock_off_sht.range("A1"))
        input_sht.range(f"A1").expand("right").copy(amt_diff_sht.range("A1"))
        input_sht.range(f"A1").expand("right").copy(diff_month_sht.range("A1"))
        ignore_check= False
        if day == 15:#replace else append

            knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
            knock_off_sht.range(f"A2:A{knock_off_last_row}").api.EntireRow.Delete()
            amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
            amt_diff_sht.range(f"A2:A{amt_diff_last_row}").api.EntireRow.Delete()
            diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row
            if diff_month_last_row!=1:
                diff_month_sht.range(f"A2:A{diff_month_last_row}").api.EntireRow.Delete()

        knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
        amt_diff_last_row = amt_diff_sht.range(f"A{amt_diff_sht.cells.last_cell.row}").end("up").row
        diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row

        i=2

        while i <=last_row:
            if not ignore_check:
                #Checking Mrn with next pjv row
                if input_sht.range(f"{voucher_col_col_letter}{i}").value.split(":")[1] == input_sht.range(f"{mrn_col_letter}{i+1}").value:
                    #Condition for knock off and amount diff tab
                    
                    if input_sht.range(f"{date_col_letter}{i}").value.month == curr_month_num:
                        #knock Off
                        if input_sht.range(f"{credit_col_letter}{i}").value is not None and input_sht.range(f"{debit_col_letter}{i+1}").value is not None:
                            i = knockOffAmtDiff(i, i+1, wb, input_sht, input_sht, credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)
                        else:#interchange debit and credit col
                            i = knockOffAmtDiff(i, i+1, wb, input_sht, input_sht, debit_col_letter, credit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)




                        
                        last_row = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
                        # ignore_check=True
                        # print("Move both enteries to knock off tab")
                    #prev month MRN Booked in Current Month
                    elif input_sht.range(f"{date_col_letter}{2}").value.month != curr_month_num:
                        print("Move both enteries to prev month MRN Booked in Current Month")
                        diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row

                        input_sht.range(f"{i}:{i+1}").api.Copy()
                        wb.activate()
                        diff_month_sht.activate()
                        diff_month_sht.range(f"A{diff_month_last_row+1}").api.Select()
                        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                        diff_month_sht.autofit()

                        # input_sht.range(f"{i}:{i+1}").copy(diff_month_sht.range(f"A{diff_month_last_row+1}"))

                        input_sht.range(f"{i}:{i+1}").delete()

                        i-=1
                    else:
                        print(f"New case for row number {i}")
                else:
                    print(f"MRN no or pjv line not found in row {i}",end="\n")
                    print(f"Keeping this row for ethanol accrual")
            else:
                print(f"pjv row num is {i}")
            i+=1
        ##########################################################################################################################
        #Filtering out remaining
        input_sht.autofit()
        input_sht.activate()
        font_colour,Interior_colour = conditional_formatting(railcar_col_letter,input_sht,wb)
        input_sht.api.AutoFilterMode=False
        input_sht.api.Range(f"{railcar_col_letter}1").AutoFilter(Field:=f"{railcar_col+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        input_sht.range(f"A1:{curr_last_col_letter}{curr_last_row}").api.Sort(Key1=input_sht.range(f"{railcar_col_letter}1:{railcar_col_letter}{curr_last_row}").api,
            Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

        #Finding filtered range

        row_range, sp_lst_row, sp_address = row_range_calc(input_sht, wb)
        curr_railcar = input_sht.range(f"{railcar_col_letter}{row_range[0]}").value
        curr_index = 0
        final_index = 0
        i=0
        # for i in row_range:
        while sp_lst_row!=1:
            if (input_sht.range(f"{railcar_col_letter}{row_range[i]}").value!=curr_railcar) or (row_range[i] == row_range[-1]):
                input_sht.activate()
                final_index = i-1
                
                if row_range[i] == row_range[-1]:
                    final_index = i
                    
                #sum of Debit amount and Credit Amount
                debit_value = input_sht.range(f"{debit_col_letter}{row_range[curr_index]}:{debit_col_letter}{row_range[final_index]}").value
                credit_value = input_sht.range(f"{credit_col_letter}{row_range[curr_index]}:{credit_col_letter}{row_range[final_index]}").value
                if isinstance(debit_value, list):
                    debit_sum = sum(filter(None, debit_value))
                else:
                    debit_sum = debit_value
                if isinstance(credit_value, list):
                    credit_sum = sum(filter(None, credit_value))
                else:
                    credit_sum = credit_value
                if (debit_sum+credit_sum) == 0:
                    # if input_sht.range(f"{credit_col_letter}{row_range[curr_index]}").value is not None and input_sht.range(f"{debit_col_letter}{row_range[final_index]}").value is not None:
                    knock_off_last_row = knock_off_sht.range(f"A{knock_off_sht.cells.last_cell.row}").end("up").row
                    input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").copy(knock_off_sht.range(f"A{knock_off_last_row+1}"))
                    input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").delete()

                    
                    #     i = knockOffAmtDiff(row_range[curr_index], row_range[final_index], wb, input_sht, credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)
                    # else:#interchange debit and credit col
                    #     i = knockOffAmtDiff(row_range[curr_index], row_range[final_index], wb, input_sht, debit_col_letter, credit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)
                    # i = knockOffAmtDiff(row_range[curr_index], row_range[final_index], wb, input_sht, credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, mrn_col_letter)
                    row_range, sp_lst_row, sp_address = row_range_calc(input_sht, wb)
                    curr_railcar = input_sht.range(f"{railcar_col_letter}{row_range[0]}").value
                    curr_index = 0
                    i=0
                else:
                    print("New condition found moving that data to Special_Sheet")
                    try:
                        spcl_sht = wb.sheets["Special_Sheet"]
                    except:
                        spcl_sht = wb.sheets.add(name="Special_Sheet", after=reco_sht)

                    input_sht.range(f"A1").expand("right").copy(spcl_sht.range("A1"))
                    spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row


                    input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").copy(spcl_sht.range(f"A{spcl_sht_last_row+1}"))

                    input_sht.range(f"{row_range[curr_index]}:{row_range[final_index]}").delete()
                    row_range, sp_lst_row, sp_address = row_range_calc(input_sht, wb)
                    curr_railcar = input_sht.range(f"{railcar_col_letter}{row_range[0]}").value
                    curr_index = 0
                    i=0

            
                # curr_index=final_index
                i-=1
            
            
            
            
            sp_lst_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
            i+=1

        #################################Add logic again copy back data from special sheet to input sheet#########################        
        input_sht.api.AutoFilterMode=False
        spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row
        last_row = input_sht.range(f"A{input_sht.cells.last_cell.row}").end("up").row
        spcl_sht.range(f"2:{spcl_sht_last_row}").copy(input_sht.range(f"A{last_row+1}"))

        #Deleting copied data from special sheet
        spcl_sht.range(f"2:{spcl_sht_last_row}").delete()

        input_sht.range(f"A1:{curr_last_col_letter}{curr_last_row}").api.Sort(Key1=input_sht.range(f"{railcar_col_letter}1:{railcar_col_letter}{curr_last_row}").api,
            Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

        #MRR will be donw at end
        #Now pjv logic

        
        input_sht.api.Range(f"{voucher_col_col_letter}1").AutoFilter(Field:=f"{voucher_col+1}", Criteria1:="Pjv*", Operator:=7)
        sp_lst_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
        try:
            pjv_sht = wb.sheets.add("PJV",after=input_sht)
        except:
            pjv_sht = wb.sheets("PJV")
        input_sht.activate()
        input_sht.api.Range(f"A1:{curr_last_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.selection.copy(pjv_sht.range(f"A1"))
        pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row

        input_sht.activate()
        input_sht.api.Range(f"A2:{curr_last_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(win32c.DeleteShiftDirection.xlShiftUp)
        input_sht.api.AutoFilterMode=False
        

        #Add MRN move to logic here

        ###Ethanol Accrual Sheet logic starts here
        eth_acr_sht = wb.sheets("Ethanol Accrual")
        eth_col_list = eth_acr_sht.range("A1").expand('right').value
        eth_credit_col = eth_col_list.index("Credit Amount")
        eth_credit_col_letter = num_to_col_letters(eth_credit_col+1)

        eth_final_amt_col = eth_col_list.index("Final Amount")
        eth_final_amt_col_letter = num_to_col_letters(eth_final_amt_col+1)

        eth_rail_col = eth_col_list.index("Rail Car/Truck #")
        eth_rail_col_letter = num_to_col_letters(eth_rail_col+1)

        eth_last_col = len(eth_col_list)
        eth_last_col_letter = num_to_col_letters(eth_last_col)
        eth_trueup_col = eth_col_list.index("TrueUp")
        eth_trueup_col_letter = num_to_col_letters(eth_trueup_col+1)
        eth_bol_col = eth_col_list.index("BOL Number")
        eth_bol_col_letter = num_to_col_letters(eth_bol_col+1)

        #filter out red color cell from credit amount column
        eth_acr_sht.api.AutoFilterMode=False
        eth_acr_sht.api.Range(f"{eth_credit_col_letter}1").AutoFilter(Field:=f"{eth_credit_col+1}", Criteria1:=Interior_colour, 
        Operator:=win32c.AutoFilterOperator.xlFilterNoFill)
        eth_acr_sht.activate()
        sp_lst_row = eth_acr_sht.range(f'A'+ str(eth_acr_sht.cells.last_cell.row)).end('up').row
        
        
        
        pjv_col_list = pjv_sht.range(f"A1").expand('right').value
        pjv_last_col = len(pjv_col_list)
        pjv_last_col_letter = num_to_col_letters(pjv_last_col+1)

        pjv_trueup_col = pjv_last_col+1
        pjv_trueup_col_letter = num_to_col_letters(pjv_trueup_col+1)

        pjv_credit_col = pjv_col_list.index("Credit Amount")
        pjv_credit_col_letter = num_to_col_letters(pjv_credit_col+1)

        pjv_debit_col = pjv_col_list.index("Debit Amount")
        pjv_debit_col_letter = num_to_col_letters(pjv_debit_col+1)

        pjv_mrn_col = pjv_col_list.index("MRN No:")
        pjv_mrn_col_letter = num_to_col_letters(pjv_mrn_col+1)
        
        pjv_railcar_col = pjv_col_list.index("Rail Car/Truck #")
        pjv_railcar_col_letter = num_to_col_letters(pjv_railcar_col+1)

        pjv_bol_col = pjv_col_list.index("BOL Number")
        pjv_bol_col_letter = num_to_col_letters(pjv_bol_col+1)

        pjv_voucher_col = pjv_col_list.index("Voucher No")
        pjv_voucher_col_letter = num_to_col_letters(pjv_voucher_col+1)

        pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row
        
        
        #Pasting BOL numbers from pjv sheet
        # pjv_sht.range(f"{pjv_bol_col_letter}2:{pjv_bol_col_letter}{pjv_last_row}").copy(eth_acr_sht.range(f"{eth_bol_col_letter}{sp_lst_row+6}"))
        #using railcar instead of bol number for getting data from ethanol accrual sheet
        pjv_sht.range(f"{pjv_railcar_col_letter}2:{pjv_railcar_col_letter}{pjv_last_row}").copy(eth_acr_sht.range(f"{eth_rail_col_letter}{sp_lst_row+6}"))

        font_colour,Interior_colour = conditional_formatting(eth_rail_col_letter,eth_acr_sht,wb)
        # eth_acr_sht.api.AutoFilterMode=False
        eth_acr_sht.api.Range(f"{eth_rail_col_letter}1").AutoFilter(Field:=f"{eth_rail_col+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)
        
        
        eth_acr_sht.api.Range(f"B1:{eth_trueup_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.selection.copy(pjv_sht.range(f"A{pjv_last_row+1}"))

        #deleting bol numbers copied from pjv sheet in eth accr sheet
        eth_acr_sht.range(f"{eth_bol_col_letter}{sp_lst_row+6}").expand("down").clear()


        #Deleting copied data from ethanol Accrual Sheet
        eth_acr_sht.api.Range(f"A2:{eth_last_col_letter}{sp_lst_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()#Delete(win32c.DeleteShiftDirection.xlShiftUp)
        wb.app.selection.delete(shift='left')
        eth_acr_sht.api.Range(f"A2:{eth_last_col_letter}{sp_lst_row}").EntireRow.SpecialCells(win32c.CellType.xlCellTypeVisible).Select()#Delete(win32c.DeleteShiftDirection.xlShiftUp)
        wb.app.selection.delete(shift='up')
        # input_sht.api.AutoFilterMode=False

############################################Update logic from above for adding bol number of pjv for duplicate check #########################################################################################################################################
        pjv_sht.activate()
        pjv_col2_list = pjv_sht.range(f"A{pjv_last_row+1}").expand('right').value
        pjv_fin_amt2_col = pjv_col2_list.index("Final Amount")
        pjv_fin_amt2_col_letter = num_to_col_letters(pjv_fin_amt2_col+1)
        pjv_trueup2_col = pjv_col2_list.index("TrueUp")
        pjv_trueup2_col_letter = num_to_col_letters(pjv_trueup2_col+1)
        pjv_credit2_col = pjv_col2_list.index("Credit Amount")
        pjv_credit2_col_letter = num_to_col_letters(pjv_credit2_col+1)

        

        
        

        # #HighLighting duplicate Railcar numbers
        # font_colour,Interior_colour = conditional_formatting(pjv_railcar_col,pjv_sht,wb)

        # pjv_sht.api.AutoFilterMode=False
        # pjv_sht.api.Range(f"{pjv_railcar_col_letter}1").AutoFilter(Field:=f"{pjv_railcar_col+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterNoFill)

        # pjv_sht.activate()
        # sp_lst_row = pjv_sht.range(f'A'+ str(pjv_sht.cells.last_cell.row)).end('up').row
        # pjv_sht.api.Range(f"A2:{pjv_last_col_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        # wb.app.selection.copy(eth_acr_sht.range(f"A{pjv_last_row+1}"))
        pjv_sht.activate()
        
        #Making Trueup Col
        pjv_sht.range(f"{pjv_trueup_col_letter}1").value = "TrueUp"

        #Deletion and column shifting logic
        pjv_sht.range(f"{pjv_fin_amt2_col_letter}{pjv_last_row+1}").expand("down").delete()
        pjv_col2_list = pjv_sht.range(f"A{pjv_last_row+1}").expand('right').value
        pjv_trueup2_col = pjv_col2_list.index("TrueUp")
        pjv_trueup2_col_letter = num_to_col_letters(pjv_trueup2_col+1)
        

        pjv_sht.range(f"{pjv_trueup2_col_letter}{pjv_last_row+1}").expand("down").api.Cut(pjv_sht.range(f"{pjv_trueup_col_letter}{pjv_last_row+1}").api)
        
        pjv_sht.range(f"{pjv_credit2_col_letter}{pjv_last_row+1}").expand("down").api.Cut(pjv_sht.range(f"{pjv_credit_col_letter}{pjv_last_row+1}").api)

        #Deleting secondary headers
        pjv_sht.range(f"{pjv_last_row+1}:{pjv_last_row+1}").delete()

        #Sorting by railcar
        pjv_last_row = pjv_sht.range(f'A'+ str(pjv_sht.cells.last_cell.row)).end('up').row
        
        pjv_sht.range(f"A1:{pjv_trueup_col_letter}{pjv_last_row}").api.Sort(Key1=pjv_sht.range(f"{pjv_voucher_col_letter}1:{pjv_voucher_col_letter}{pjv_last_row}").api,
        Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)


        pjv_sht.range(f"A1:{pjv_trueup_col_letter}{pjv_last_row}").api.Sort(Key1=pjv_sht.range(f"{pjv_railcar_col_letter}1:{pjv_railcar_col_letter}{pjv_last_row}").api,
        Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

        i=2
        while i <=pjv_last_row:
            if not ignore_check:
                #Checking Mrn with next pjv row
                if pjv_sht.range(f"{voucher_col_col_letter}{i}").value.split(":")[1] == pjv_sht.range(f"{mrn_col_letter}{i+1}").value:
                    #Condition for knock off and amount diff tab
                    
                    if pjv_sht.range(f"{date_col_letter}{i}").value.month == curr_month_num:
                        #knock Off
                        if pjv_sht.range(f"{pjv_credit_col_letter}{i}").value is not None and pjv_sht.range(f"{debit_col_letter}{i+1}").value is not None:
                            i = knockOffAmtDiff(i, i+1, wb, pjv_sht, pjv_sht, pjv_credit_col_letter,debit_col_letter, knock_off_sht, amt_diff_sht, pjv_mrn_col_letter)
                        else:#interchange debit and credit col
                            i = knockOffAmtDiff(i, i+1, wb, pjv_sht, pjv_sht, pjv_debit_col_letter, pjv_credit_col_letter, knock_off_sht, amt_diff_sht, pjv_mrn_col_letter)




                        
                        pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row
                        # ignore_check=True
                        # print("Move both enteries to knock off tab")
                    #prev month MRN Booked in Current Month
                    elif pjv_sht.range(f"{date_col_letter}{2}").value.month != curr_month_num:
                        print("Move both enteries to prev month MRN Booked in Current Month")
                        diff_month_last_row = diff_month_sht.range(f"A{diff_month_sht.cells.last_cell.row}").end("up").row

                        pjv_sht.range(f"{i}:{i+1}").api.Copy()
                        wb.activate()
                        diff_month_sht.activate()
                        diff_month_sht.range(f"A{diff_month_last_row+1}").api.Select()
                        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                        diff_month_sht.autofit()

                        # pjv_sht.range(f"{i}:{i+1}").copy(diff_month_sht.range(f"A{diff_month_last_row+1}"))

                        pjv_sht.range(f"{i}:{i+1}").delete()

                        i-=1
                    else:
                        print(f"New case for row number {i}")
                else:
                    print(f"MRN no or pjv line not found in row {i}",end="\n")
                    print(f"Keeping this row for ethanol accrual")
            else:
                print(f"pjv row num is {i}")
            i+=1
            pjv_last_row = pjv_sht.range(f"A{pjv_sht.cells.last_cell.row}").end("up").row

        
        pjv_sht.range(f"A1").expand("right").copy(spcl_sht.range("A1"))
        try:
            spcl_sht = wb.sheets["Special_Sheet"]
        except:
            spcl_sht = wb.sheets.add(name="Special_Sheet", after=reco_sht)
        spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row


        pjv_sht.range(f"2:{pjv_last_row}").copy(spcl_sht.range(f"A{spcl_sht_last_row+1}"))

        pjv_sht.range(f"2:{pjv_last_row}").delete()


        #Now deleting pjv Sheet
        pjv_sht.delete()

        #Now checking input sheet for remaing rows
        input_sht.activate()
        #Removing MRR Logic
        input_sht.api.AutoFilterMode=False
        input_sht.api.Range(f"{voucher_col_col_letter}1").AutoFilter(Field:=f"{voucher_col+1}", Criteria1:="MRR*", Operator:=7)

        #searching all bol numbers in ethanol accrual sheet for each mrr found in inpurt sheet
        row_range, sp_lst_row, sp_address = row_range_calc(input_sht, wb)
        curr=0
        for i in range(len(row_range)):

            bol_num = input_sht.range(f"{bol_col_letter}{row_range[i]}").value
            eth_acr_sht.activate()
            eth_acr_sht.api.AutoFilterMode=False
            try:
                eth_acr_sht.api.Cells.Find(What:=bol_num , After:=eth_acr_sht.api.Application.ActiveCell, LookIn:=win32c.FindLookIn.xlFormulas,
                LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()

                cell_value = eth_acr_sht.api.Application.ActiveCell.Address.replace("$","")
                row_num = int(re.findall(r'\d+', cell_value)[0])

                #Copy delete logic
                curr=knockOffAmtDiff(row_range[i],row_num, wb, input_sht, eth_acr_sht, debit_col_letter, eth_credit_col_letter, knock_off_sht, amt_diff_sht,
                mrn_col_letter, eth_trueup_col_letter)
                curr = row_range[i]-curr
                

            except:
                spcl_sht_last_row = spcl_sht.range(f"A{spcl_sht.cells.last_cell.row}").end("up").row
                input_sht.range(f"{row_range[i]-curr}:{row_range[i]-curr}").copy(spcl_sht.range(f"A{spcl_sht_last_row+1}"))
            
                input_sht.range(f"{row_range[i]-curr}:{row_range[i]-curr}").delete()


        #Logic for moving remaining mrn in input sheet to ethanol accrual sheet
        input_sht.api.AutoFilterMode=False
        curr_last_row = input_sht.range(f'A'+ str(input_sht.cells.last_cell.row)).end('up').row
        eth_last_row = eth_acr_sht.range(f'A'+ str(eth_acr_sht.cells.last_cell.row)).end('up').row

        input_sht.activate()
        row_count = input_sht.range(f"A2").expand("down").count
        for i in range(0,row_count):
            eth_acr_sht.api.Range(f"B{eth_last_row+1}").EntireRow.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
        input_sht.range(f"A2:{credit_col_letter}{curr_last_row}").copy(eth_acr_sht.range(f"B{eth_last_row+1}"))
        input_sht.range(f"A2:{credit_col_letter}{curr_last_row}").api.EntireRow.Delete()
        eth_acr_sht.range(f"M{eth_last_row+1}").expand("down").copy(eth_acr_sht.range(f"L{eth_last_row+1}"))
        # eth_acr_sht.range(f"M{eth_last_row+1}").expand("down").clear()
        eth_acr_sht.range(f"M{eth_last_row+1}").expand("down").api.NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        eth_acr_sht.range(f"L{eth_last_row+1}").expand("down").api.NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        
        eth_acr_sht.activate()
        #Refreshing pivot table in ethanol accrual tab
        pivotCount = wb.api.ActiveSheet.PivotTables().Count
         # 'INPUT DATA'!$A$3:$I$86
        for j in range(1, pivotCount+1):     
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

        wb.save(output_location+f"\\Open GR {month}{day}.xlsx")


        print("Done")
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass



# openGr("12.31.2022", "12.31.2022")