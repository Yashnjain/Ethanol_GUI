import time
import os,re
import xlwings as xw
import xlwings.constants as win32c


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
            
def ar_ageing_master(input_date, output_date):
    try:
        # today_date=date.today()     
        job_name = 'Ar_Ageing_Master'
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
        year = input_date.split(".")[-1]
        bulk_input_sheet= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_bulk\Output'+f'\\AR Aging Bulk {month}{day}-updated.xlsx'
        rack_input_sheet= r'J:\India\BBR\IT_BBR\Reports\Ar Ageing_rack\Output'+f'\\AR Aging Rack {month}{day}-updated.xlsx'

        temp_location = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing Master\Template'+f'\\Ar Ageing Master Template.xlsx'

        output_location = r'J:\India\BBR\IT_BBR\Reports\Ar Ageing Master\Output'
        if not os.path.exists(bulk_input_sheet):
            return(f"{bulk_input_sheet} Excel file not present for date {input_date}")  
        if not os.path.exists(rack_input_sheet):
            return(f"{rack_input_sheet} Excel file not present for date {input_date}")  
        if not os.path.exists(temp_location):
            return(f"{temp_location} Excel file not present,Please contact IT") 
                                      
        retry=0
        while retry < 10:
            try:
                bulk_wb = xw.Book(bulk_input_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e                     
        retry=0
        while retry < 10:
            try:
                rack_wb = xw.Book(rack_input_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        # bulk_input_tab = bulk_wb.sheets[f'Bulk_Data(IT)']
        updated_bulk_tab = bulk_wb.sheets[f'Updated_Data(IT)']
        bulk_output_tab_1 = bulk_wb.sheets[f'Bulk_Data(IT)']               
        bulk_output_tab = bulk_wb.sheets[f'Bulk_Data(IT)(2)']
        rack_output_tab = rack_wb.sheets[f'Rack_Data(IT)']       

        retry=0
        while retry < 10:
            try:
                wb = xw.Book(temp_location) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        initial_tab= wb.sheets[0]
        updated_bulk_tab.api.Copy(Before=wb.api.Sheets(1))
        bulk_output_tab_1.api.Copy(Before=wb.api.Sheets(2))
        bulk_output_tab.api.Copy(Before=wb.api.Sheets(3))
        rack_output_tab.api.Copy(Before=wb.api.Sheets(4))
        # wb.sheets.add("Merged_Results(IT)",after=input_tab)
        initial_tab.name = "Merged_Results(IT)"
        
        bulk_tab= wb.sheets[f'Bulk_Data(IT)(2)']
        rack_tab= wb.sheets[f'Rack_Data(IT)']

        bulk_tab.activate()

        bulk_tab.range(f"B8").select()
        bulk_add = bulk_tab.range(f"B8").end('down').address.replace("$",'')
        bulk_add_row = re.findall("\d+",bulk_add)[0]

        bulk_tab.api.Range(f"B8:D{bulk_add_row}").Copy() 
        initial_tab.api.Range(f"B2")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)

        #####formatting the merged sheet
        ###for bulk
        initial_tab.activate()
        initial_tab.range(f"A2").value = 'Bulk'

        customer_last_row = initial_tab.range(f'B'+ str(initial_tab.cells.last_cell.row)).end('up').row

        initial_tab.api.Range(f"A2:A{customer_last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()

        initial_tab.range(f"F2").value = "=+XLOOKUP(B2,'J:\India\Hamilton\Temporary\BBR Working\[BBR Master.xlsx]Main'!$B:$B,'J:\India\Hamilton\Temporary\BBR Working\[BBR Master.xlsx]Main'!$D:$D,0)"
        initial_tab.api.Range(f"F2:F{customer_last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()

        ####for rack
        rackini_row_no = customer_last_row + 10
        initial_tab.range(f"A{rackini_row_no}").value = 'Rack'

        rack_tab.activate()
        rack_tab.range(f"B8").select()
        rack_add = rack_tab.range(f"B8").end('down').address.replace("$",'')
        rack_add_row = re.findall("\d+",rack_add)[0]

        rack_tab.api.Range(f"B8:D{rack_add_row}").Copy() 
        initial_tab.api.Range(f"B{rackini_row_no}")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)

        customer_last_row_rack = initial_tab.range(f'B'+ str(initial_tab.cells.last_cell.row)).end('up').row

        time.sleep(1)
        initial_tab.activate()
        initial_tab.api.Range(f"A{rackini_row_no}:A{customer_last_row_rack}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()        

        initial_tab.range(f"F{rackini_row_no}").value = "=+XLOOKUP(B2,'J:\India\Hamilton\Temporary\BBR Working\[BBR Master.xlsx]Main'!$B:$B,'J:\India\Hamilton\Temporary\BBR Working\[BBR Master.xlsx]Main'!$D:$D,0)"
        initial_tab.api.Range(f"F{rackini_row_no}:F{customer_last_row_rack}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()

        ####total section #########
        total_row_no = customer_last_row_rack + 5

        initial_tab.range(f"C{total_row_no}").value = f"=SUM(C2:C{total_row_no-1})"
        initial_tab.range(f"D{total_row_no}").value = f"=SUM(D2:D{total_row_no-1})"

        ####finding totals for bulk and rack ######
        ##bulk first
        bulk_tab.activate()
        bulk_tab.range(f"A1").select()
        bulk_tab.api.Cells.Find(What:="Total", After:=bulk_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        cell_value_bulk = bulk_tab.api.Application.ActiveCell.Address.replace("$","")
        row_value_bulk = re.findall("\d+",cell_value_bulk)[0]

        ##rack second
        rack_tab.activate()
        rack_tab.range(f"A1").select()
        rack_tab.api.Cells.Find(What:="Total", After:=rack_tab.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        cell_value_rack = rack_tab.api.Application.ActiveCell.Address.replace("$","")
        row_value_rack = re.findall("\d+",cell_value_rack)[0]        

        initial_tab.activate()
        initial_tab.range(f"C{total_row_no+2}").value =f"=+'Bulk_Data(IT)(2)'!C{row_value_bulk}+'Rack_Data(IT)'!C{row_value_rack}"
        initial_tab.range(f"D{total_row_no+2}").value =f"=+'Bulk_Data(IT)(2)'!D{row_value_bulk}+'Rack_Data(IT)'!D{row_value_rack}"


        ###finding difference(if any)

        initial_tab.range(f"C{total_row_no+4}").value = f"=+C{total_row_no}-C{total_row_no+2}"
        initial_tab.range(f"D{total_row_no+4}").value = f"=+D{total_row_no}-D{total_row_no+2}"

        insert_all_borders(cellrange=f"C{total_row_no+4}",working_sheet=initial_tab,working_workbook=wb)
        insert_all_borders(cellrange=f"D{total_row_no+4}",working_sheet=initial_tab,working_workbook=wb)

        interior_coloring(colour_value=13311,cellrange=f"C{total_row_no+4}",working_sheet=initial_tab,working_workbook=wb)
        interior_coloring(colour_value=65535,cellrange=f"D{total_row_no+4}",working_sheet=initial_tab,working_workbook=wb)

        tablist={initial_tab:win32c.ThemeColor.xlThemeColorAccent2}
        for tab,color in tablist.items():
                tab.activate()
                tab.api.Tab.ThemeColor = color
                tab.autofit()
                tab.range(f"A1").select()
        initial_tab.activate()
        initial_tab.range(f"A1").select()
        initial_tab.autofit()



        wb.save(f"{output_location}\\AR Aging Master {month}{day}"+'.xlsx') 
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

