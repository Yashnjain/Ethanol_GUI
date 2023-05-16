import time
import os,re
import xlwings as xw
from datetime import datetime
import xlwings.constants as win32c
from Common.common import freezepanes_for_tab,num_to_col_letters
      


def del_car(input_date, output_date):
    try:     
        job_name = 'Delivered_Cars_Automation'
        check_del = None
        check_acc = None
        ws2 = None
        ws3 = None
        month = input_date.split(".")[0]
        day = input_date.split(".")[1]
        year = input_date.split(".")[2]
        dt = datetime.strptime(input_date,"%m.%d.%Y")
        input_sheet= r'J:\India\BBR\IT_BBR\Reports\Delivered Cars Working\Input'+f'\\DEL_{year}{month}{day}.xlsx'
        output_location = r'J:\India\BBR\IT_BBR\Reports\Delivered Cars Working\Output' 
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
        arrival_date_column_no = column_list.index('Arrival Date')+1
        arrival_date_column_letter=num_to_col_letters(arrival_date_column_no)
        inco_column_no = column_list.index('Inco terms')+1
        inco_column_letter=num_to_col_letters(inco_column_no)        
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        dict1={"=":[bldate_No_column_no,bldate_No_column_letter,"B"],f">{datetime.strptime(input_date,'%m.%d.%Y')}":[bldate_No_column_no,bldate_No_column_letter,"A"],f"<={datetime.strptime(input_date,'%m.%d.%Y')}":[arrival_date_column_no,arrival_date_column_letter,"A"]}
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

        input_tab.api.Range(f"{inco_column_letter}1").AutoFilter(Field:=f'{inco_column_no}', Criteria1:=["DEL"])
        sp_lst_row = input_tab.range(f'{inco_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"{inco_column_letter}2:{inco_column_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row
        if int(sp_lst_row)!=1:
            wb.sheets.add("DEL(IT)",after=input_tab)
            ###logger.info("Clearing contents for new sheet")
            wb.sheets["DEL(IT)"].clear_contents()
            ws2=wb.sheets["DEL(IT)"]
            input_tab.activate()  
            input_tab.range(f"A1:{last_column_letter}{sp_lst_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
            ws2.api.Range(f"A1")._PasteSpecial(Paste=13)
            wb.app.api.CutCopyMode=False
            input_tab.api.Range(f"{sp_initial_rw}:{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
            time.sleep(1)
            wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            time.sleep(1)
            check_del=True
        input_tab.api.AutoFilterMode=False 

        if check_del:
            ws2.activate()

            db_No_column_no = column_list.index('Date -Base')+1
            db_No_column_letter=num_to_col_letters(db_No_column_no)

            ws2.api.Range(f"{db_No_column_letter}1").AutoFilter(Field:=f'{db_No_column_no}', Criteria1:=[f">{datetime.strptime(input_date,'%m.%d.%Y')}"])
            sp_lst_row = ws2.range(f'{db_No_column_letter}'+ str(ws2.cells.last_cell.row)).end('up').row
            sp_address= ws2.api.Range(f"{db_No_column_letter}2:{db_No_column_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            lst_rw = ws2.range('A'+ str(ws2.cells.last_cell.row)).end('up').row
            if int(sp_lst_row)!=1: 
                ws2.api.Range(f"{sp_initial_rw}:{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                time.sleep(1)
                wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                time.sleep(1)
            ws2.api.AutoFilterMode=False 

        input_tab.api.Range(f"{db_No_column_letter}1").AutoFilter(Field:=f'{db_No_column_no}', Criteria1:=[f">{datetime.strptime(input_date,'%m.%d.%Y')}"])
        sp_lst_row = input_tab.range(f'{db_No_column_letter}'+ str(input_tab.cells.last_cell.row)).end('up').row
        sp_address= input_tab.api.Range(f"{db_No_column_letter}2:{db_No_column_letter}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        lst_rw = input_tab.range('A'+ str(input_tab.cells.last_cell.row)).end('up').row
        if int(sp_lst_row)!=1:
            wb.sheets.add("Accrued MRN(IT)",after=ws2)
            ###logger.info("Clearing contents for new sheet")
            wb.sheets["Accrued MRN(IT)"].clear_contents()
            ws3=wb.sheets["Accrued MRN(IT)"]
            input_tab.activate()  
            input_tab.range(f"A1:{last_column_letter}{sp_lst_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
            ws3.api.Range(f"A1")._PasteSpecial(Paste=13)
            wb.app.api.CutCopyMode=False
            check_acc = True
        input_tab.api.AutoFilterMode=False
    	
        
        if check_acc:
            ws3.activate() 
            lst_rw = ws3.range('A'+ str(ws3.cells.last_cell.row)).end('up').row
            ###logger.info("Declaring Variables for columns and rows")
            last_column = ws3.range('A1').end('right').last_cell.column
            last_column_letter=num_to_col_letters(ws3.range('A1').end('right').last_cell.column)
            ###logger.info("Creating Pivot Table")
            PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'{ws3.name}\'!R1C1:R{lst_rw}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
            PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'{ws3.name}'!R{lst_rw+5}C3", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
            PivotTable.PivotFields('Voucher No.-Base ').Orientation = win32c.PivotFieldOrientation.xlRowField
            PivotTable.PivotFields('Voucher No.-Base ').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
            PivotTable.PivotFields('Vendor Ref').Orientation = win32c.PivotFieldOrientation.xlRowField
            PivotTable.PivotFields('Vendor Ref').Position = 1
            PivotTable.PivotFields('Voucher No.-Base ').Position = 2
            PivotTable.PivotFields('Vendor Ref').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
            PivotTable.PivotFields('Amount').Orientation = win32c.PivotFieldOrientation.xlDataField
            # PivotTable.PivotFields('Tax').Orientation = win32c.PivotFieldOrientation.xlDataField
            PivotTable.RowAxisLayout(1)
            wb.api.ActiveSheet.PivotTables("PivotTable1").PivotFields('Vendor Ref').RepeatLabels = True
            time.sleep(1)

        
        try:
            tablist=[input_tab,ws2,ws3]
            for tab in tablist:
                if tab is not None:
                    freezepanes_for_tab(cellrange="2:2",working_sheet=tab,working_workbook=wb) 

            tablist={input_tab:win32c.ThemeColor.xlThemeColorAccent2,ws2:win32c.ThemeColor.xlThemeColorAccent6,ws3:win32c.ThemeColor.xlThemeColorAccent4}
            for tab,color in tablist.items():
                    if tab is not None:
                        tab.activate()
                        tab.api.Tab.ThemeColor = color
                        tab.autofit()
                        tab.range(f"A1").select()
            input_tab.activate()
        except:
            pass    

        wb.save(f"{output_location}\\DEL_{year}{month}{day} - updated"+'.xlsx') 
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
