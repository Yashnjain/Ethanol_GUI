import xlwings as xw
import xlwings.constants as win32c
import os, time
from datetime import datetime, timedelta
from tabula import read_pdf
import pandas as pd
import glob
import re
from tkinter import messagebox
from Common.common import last_day_of_month,row_range_calc, conditional_formatting_uniq, \
    num_to_col_letters, mrn_pdf_extractor, rack_pdf_data_extractor



def rackbacktrack(input_date, output_date):
    try:
        input_datetime = datetime.strptime(input_date, "%m.%d.%Y")
        file_date = datetime.strftime(input_datetime, "%Y%m%d")

        input_month = datetime.strftime(input_datetime, "%m")
        input_day = input_datetime.day
        input_year = input_datetime.year
        input_year2 = datetime.strftime(input_datetime, "%y")
        j_loc = r"J:\India\BBR\IT_BBR\Reports"
        j_loc_bbr = f"J:\\India\\BBR\\{input_year}\\BBR_{file_date}"
        # curr_loc = os.getcwd()
        # input_sheet= curr_loc+r'\RackBackTrack\Raw Files'+f'\\Purchase by vendor (Back track) for Cost - {file_date}.xlsx'
        input_sheet= j_loc+r'\RackBackTrack\Raw Files'+f'\\Purchase by vendor (Back track) for Cost - {file_date}.xlsx'
        output_location = j_loc+r'\RackBackTrack\Output Files'
        # output_location = curr_loc+r'\RackBackTrack\Output Files' 
        # rack_pdf = j_loc+r'\RackBackTrack\Raw Files'+f'\\Rack.pdf'
        # rack_pdf = curr_loc+r'\RackBackTrack\Raw Files'+f'\\Rack.pdf'
        rack_pdf = f"{j_loc_bbr}\\Rack.pdf"
        input_cta = j_loc_bbr+"\\Rack MTM.xls"
        # input_cta = j_loc+r'\RackBackTrack\Raw Files'+f"\\BioUrjaNet.xls"
        # input_cta = curr_loc+r'\RackBackTrack\Raw Files'+f"\\Rack MTM.xls"

        mrn_pdf_loc = f"J:\\AP - RACK\\AP Rack {input_year}\\Rack Purchase Daily\\{input_month}-{input_year}"
        monthly_mrn_loc = f"{j_loc_bbr}\\Rack Inventory.pdf"

        input_day_month = datetime.strftime(input_datetime-timedelta(days=0), "%b%d")
        input_month_year = datetime.strftime(input_datetime-timedelta(days=0), "%b %Y")
        input_day_month_other = datetime.strftime(input_datetime-timedelta(days=1), "%b%d")
        other_loc = r'S:\Magellan Setup\Pricing\_Price Changes'+f"\\Daily Pricing Template -{input_day_month_other}.xlsx"

        # input_open_gr= curr_loc+r'\RackBackTrack\Raw Files'+f'\\Open GR Rack.xlsx'
        input_open_gr= j_loc_bbr+f'\\Open GR Rack.xlsx'

        # input_lrti_xl = f'\\\\BIO-INDIA-FS\\India Sync$\\India\\{input_year}\\{input_month}-{input_year2}\\Little Rock Tank Inv Reco.xlsx'
        input_lrti_xl = f'\\\\BIO-INDIA-FS\India Sync$\\India\\{input_year}\\{input_month}-{input_year2}\\Little Rock Tank.xlsx'
        # input_lrti_xl = f'\\\\BIO-INDIA-FS\India Sync$\\India\\{input_year}\\{input_month}-{input_year2}\\Transfered\\Little Rock Tank Inv Reco.xlsx'
        # input_lrti_xl = curr_loc+r'\RackBackTrack\Raw Files'+f'\\Little Rock Tank Inv Reco.xlsx'


        truefile_loc = f"J:\\India\Trueup\\TrueupAutomation\\AP_Rack_TrueUp\\Output\\Rack AP Data {input_month_year}.xlsx"
        rack_po_loc = f'J:\\India\\Trueup\\TrueupAutomation\\AP_Rack_TrueUp\\Rack PO details\\{input_month+str(input_year)} AP PO.xlsx'

        last_date = datetime.strftime(last_day_of_month(input_datetime.date()), "%m.%d.%Y")
        if input_date==last_date:#Montlhy True up condition
            if not os.path.exists(truefile_loc):
                return(f"{truefile_loc} Excel file not present for date {input_date}")
            if not os.path.exists(truefile_loc):
                return(f"{rack_po_loc} Excel file not present for date {input_date}")

        retry=0
        while retry < 4:
            try:
                spcl_loc_df = pd.read_excel(other_loc)
                break
            except Exception as e:
                retry+=1
                input_day_month = datetime.strftime(input_datetime-timedelta(days=retry), "%b%d")
                other_loc = r'S:\Magellan Setup\Pricing\_Price Changes'+f"\\Daily Pricing Template -{input_day_month}.xlsx"
                
                if retry ==3:
                    other_loc = r'S:\Magellan Setup\Pricing\_Price Changes'+f"\\{input_year}\\Daily Pricing Template -{input_day_month_other}.xlsx"
                    retry=0
                    while retry < 4:
                        try:
                            spcl_loc_df = pd.read_excel(other_loc)
                            break
                        except Exception as e:
                            retry+=1
                            input_day_month = datetime.strftime(input_datetime-timedelta(days=retry), "%b%d")
                            other_loc = r'S:\Magellan Setup\Pricing\_Price Changes'+f"\\{input_year}\\Daily Pricing Template -{input_day_month}.xlsx"
                            
                            if retry ==4:
                                return(f"{other_loc} Excel file not present for date {input_day_month}")

        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")

        if not os.path.exists(rack_pdf):
            return(f"{rack_pdf} PDF file not present for date {input_date}")
        
        if not os.path.exists(input_cta):
            return(f"{input_cta} Excel file not present")

        if not os.path.exists(input_lrti_xl):
            input_lrti_xl = f'\\\\BIO-INDIA-FS\India Sync$\\India\\{input_year}\\{input_month}-{input_year2}\\Transfered\\Little Rock Tank.xlsx'
            if not os.path.exists(input_lrti_xl):
                return(f"{input_lrti_xl} Excel file not present")

        mtm_df = pd.read_excel(input_cta,sheet_name="BioUrjaNet", header=2)
        try:
            df, pdf_date = rack_pdf_data_extractor(rack_pdf)
        except Exception as e:
            if messagebox.askyesno("Error in Reading Rack pdf",f'Do you want to continue rest process?'):
                pass
            else:
                raise e
        # if pdf_date != file_date:
        #     return f"Pdf file date({pdf_date}) and and excel date({file_date} does not match please check pdf file and run job again)"  
        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet, update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        sht6 = wb.sheets("Sheet6")
        wb.activate()
        sht6.activate()
        #clearing old table data
        sht6.range(f"I4").expand('table').clear()
        sht6.range(f"I4").options(pd.DataFrame,header=None,
                                            index=False, 
                                            expand='right').value = df
        sht6.api.AutoFilterMode=False
        sht6_last_row = sht6.range(f"A{sht6.cells.last_cell.row}").end("up").row
        sht6.range(f"A1:K{sht6_last_row}").api.AutoFilter(Field:=4, Criteria1:="<0")
        if sht6.range(f'A'+ str(sht6.cells.last_cell.row)).end('up').row != 1:
            #Pop yes no logic to be added
            
            data_list = row_range_calc("D", sht6, wb)
            row_list = data_list[0]
            for row in row_list:
                if messagebox.askyesno("Negative Active Tank Found",f'Do you want this entry {row} to be neutralized from Tank Bottom'):
                    d_value = sht6.range(f'D{row}').value * -1
                    sht6.range(f"D{row}").formula = f"={sht6.range(f'D{row}').value} + {d_value}"#=-23+23
                    sht6.range(f"E{row}").formula = f"={sht6.range(f'E{row}').value} - {d_value}"#=-23+23

        inv_mtm_sht = wb.sheets("Inv. per MTM")
        wb.activate()
        inv_mtm_sht.activate()
        data_list = []
        inv_mtm_st_row = inv_mtm_sht.range(f"B1").end("down").row + 1 #+1 for excluding RowLabel Header
        # inv_mtm_last_row = inv_mtm_sht.range(f'A'+ str(inv_mtm_sht.cells.last_cell.row)).end('up').row
        loc_list = inv_mtm_sht.range(f"B{inv_mtm_st_row}").expand("down").value
        loc_dict = {"FORT SMITH, AR":"FT. SMITH", "MPLS/ST. PAUL, MN":"MPLS/ST.PAUL", "NLITTLEROCKNORTH, AR":"LITTLE ROCK", "NLITTLEROCKSOUTH, AR":"LITTLE ROCK"}
        spcl_dict = {'ODESSA, TX':'Odessa TX Magellan', 'East Houston, TX':'Houston TX Magellan'}
        for location in range(len(loc_list)-1):#Ifgnoring grand total
            if loc_list[location] in loc_dict.keys():
                index_list=mtm_df.loc[mtm_df['Unnamed: 0']==loc_dict[loc_list[location]].split(',')[0]].index.values
            elif loc_list[location] in spcl_dict.keys():
                index_list = []
                column_1 = spcl_loc_df.columns._values[0]
                ethrin_value = spcl_loc_df.iloc[spcl_loc_df.loc[spcl_loc_df[column_1]==spcl_dict[loc_list[location]]].index.values[0]+1]['Unnamed: 3']
                inv_mtm_sht.range(f"F{inv_mtm_st_row+location}").value = ethrin_value

            else:
                index_list=mtm_df.loc[mtm_df['Unnamed: 0']==loc_list[location].split(',')[0]].index.values
            #Removing extra values
            inv_mtm_sht.range(f"C{inv_mtm_st_row+location}").formula = inv_mtm_sht.range(f"C{inv_mtm_st_row+location}").formula.split(")+")[0]
            if len(index_list):
                for i in range(len(index_list)):
                    if index_list[i] != index_list[-1]:
                        filt_df = mtm_df[index_list[i]:index_list[i+1]]
                    else:
                        filt_df = mtm_df[index_list[i]:]
                    substring = 'ETHRIN'
                    ethrin_df = filt_df[filt_df.apply(lambda row: row.astype(str).str.contains(substring, case=False).any(), axis=1)]
                    if len(ethrin_df):
                        rack_df = filt_df.loc[filt_df['Unnamed: 0']=="RACK AVERAGE"]
                        rack_col = list(ethrin_df.apply(lambda row: row[row == substring].index, axis=1))[0][0]
                        ethrin_value = rack_df[rack_col]._values[0]
                        # data_list.append(ethrin_value)
                        inv_mtm_sht.range(f"F{inv_mtm_st_row+location}").value = ethrin_value
                        break
                    print(i)

        # Updating Sheet3 for Mid Month Case
        sht3 = wb.sheets("Sheet3")
        pivot_sht = wb.sheets("PIVOT")
        if input_day <=15:
            mrn_dict = {}
            date_list = []
            try:
                for pdf_file in glob.glob(mrn_pdf_loc+"\\*.pdf"):
                    filename = pdf_file.split("\\")[-1]
                    pdf_file_date = filename.replace("MRN.", "").replace(" done.pdf","").replace(".pdf", "")
                    pdf_file_date = datetime.strptime(pdf_file_date, "%m.%d.%Y").day
                    if pdf_file_date <=17:
                        #extract data from pdf
                        date_list, mrn_dict = mrn_pdf_extractor(pdf_file, mrn_dict, date_list)
            except Exception as e:
                if messagebox.askyesno("Error in Reading MRN pdf",f'Do you want to continue rest process?'):
                    pass
                else:
                    raise e

            #Filling Data in Excel
            wb.activate()
            sht3.activate()
            pivot_last_row = pivot_sht.range(f"A{pivot_sht.cells.last_cell.row}").end("up").row
            p_loc_r1 = pivot_sht.range(f"A{pivot_last_row}").end('up').row
            # for loc in range()
            if len(date_list):
                date_list = sorted(date_list)
                sht3.range(f"D1").value = date_list
                sht3_c_last_row = sht3.range(f"C{sht3.cells.last_cell.row}").end("up").row
                sht3_a_last_row = sht3.range(f"A{sht3.cells.last_cell.row}").end("up").row
                sht3_last_col_num  = len(date_list)+2
                sht3_last_col = num_to_col_letters(sht3_last_col_num)
                #Clearing old data
                sht3.range(f"D2:{sht3_last_col}{sht3_c_last_row}").clear_contents()

                s_mrn_dict = sorted(mrn_dict)
                key_num = 0
                col_num= 0 
                
                for i in range(2, sht3_a_last_row+1):
                    print(sht3.range(f"A{i}").value)
                    if sht3.range(f"A{i}").value == s_mrn_dict[key_num]:
                        col_num= 0 
                        # for date_col in range(len(date_list)):
                        #     if date_list[date_col].day<16:
                        #         if mrn_dict[s_mrn_dict[key_num]][col_num][0] == date_list[date_col]:
                        #             # date_column = num_to_col_letters(date_col+3)
                        #             date_column = num_to_col_letters(date_col+4)
                        #             print(date_column)
                        #             print("Entering value now")
                        #             sht3.range(f"{date_column}{i}").value = mrn_dict[s_mrn_dict[key_num]][col_num][1]
                        #             if col_num != len(mrn_dict[s_mrn_dict[key_num]])-1:
                        #                 col_num+=1
                        xl_date_list = sht3.range(f"D1").expand('right').value
                        for date_col in mrn_dict[s_mrn_dict[key_num]].keys():
                            if date_col.day<16:
                                date_index = xl_date_list.index(date_col)
                                date_column = num_to_col_letters(date_index+4)
                                sht3.range(f"{date_column}{i}").value = mrn_dict[s_mrn_dict[key_num]][date_col]

                        key_num+=1


             #Sheet 3 total column logic
            sht3_total_col = num_to_col_letters(sht3_last_col_num+1)
            sht3.range(f"C1").value = "Total"
            sht3.range(f"C2").formula = f"=SUM(D2:{sht3_total_col}2)"
            wb.activate()
            sht3.activate()
            sht3.autofit()
            sht3.range(f"C2:C{sht3_a_last_row}").select()
            wb.app.selection.api.FillDown()

            #Making dataframe
            sht3_df = sht3.range(f"B1:{sht3_total_col}{sht3_a_last_row}").options(pd.DataFrame, 
                                    header=1,
                                    index=False 
                                    ).value
            sht3_df = sht3_df[["No", "Total"]]
            sht3_dict = sht3_df.set_index("No").to_dict()['Total']
        else:#Update Pivot sheet for month end case
            if not os.path.exists(monthly_mrn_loc):
                return(f"{monthly_mrn_loc} Pdf file not present for date {input_date}")

            mrn_dict = {}
            date_list = []
            try:
                date_list, mrn_dict = mrn_pdf_extractor(monthly_mrn_loc, mrn_dict, date_list, rack=True)
            except Exception as e:
                if messagebox.askyesno("Error in Reading Rack MRN pdf",f'Do you want to continue rest process?'):
                    pass
                else:
                    raise e

            wb.activate()
            pivot_sht.activate()
            pivot_last_row = pivot_sht.range(f"A{pivot_sht.cells.last_cell.row}").end("up").row
            p_loc_r1 = pivot_sht.range(f"A{pivot_last_row}").end('up').row
            # p_loc_list = pivot_sht.range(f"A{p_loc_r1}").expand('down').value
            p_loc_list = pivot_sht.range(f"B{p_loc_r1}").expand('down').value

            init_row = p_loc_r1
            # for location in range(len(p_loc_list)):
            for location in p_loc_list:
                location = int(location)
                try:
                    # pivot_sht.range(f"C{location+p_loc_r1}").value = mrn_dict[p_loc_list[location].split(",")[0]]
                    pivot_sht.range(f"C{init_row}").value = mrn_dict[str(location)]
                except:
                    pivot_sht.range(f"C{init_row}").value = 0
                    # print(f"NO data found for {p_loc_list[location]}")
                init_row += 1

            #Making dataframe
            pivot_df = pivot_sht.range(f"A{p_loc_r1}:C{pivot_last_row}").options(pd.DataFrame, 
                                    header=False,
                                    index=False 
                                    ).value

            # pivot_dict = pivot_df.set_index(0).to_dict()[1]
            pivot_dict = pivot_df.set_index(0).to_dict()[2]



        #For monthly caseMultiplying by 42 to convert into gallons from barrels
        # sht3_df["Total"] = sht3_df["Total"]*42
        
        # #######Open GR Sheet 1 Logic###########################
        retry=0
        while retry < 10:
            try:
                gr_wb = xw.Book(input_open_gr, update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        try:
            wb.sheets["Sheet1"].delete()
        except:
            pass

        # #Copy Sheet1 from open gr sheet1
        gr_wb.sheets("Sheet1").copy(name="Sheet1", after=wb.sheets["Little Rock Costing"])
        open_gr_sht = wb.sheets["Sheet1"]
        gr_head_row = open_gr_sht.range("B1").end('down').end('down').row
        gr_wb.close()
        # #Unmerging Top 3 cells
        open_gr_sht.range(f"1:{gr_head_row-1}").unmerge()
        open_gr_sht.range(f"8:8").unmerge()
        gr_col_list = open_gr_sht.range(f"B{gr_head_row}").expand("right").value
        gr_last_col = len(gr_col_list)
        gr_last_col_letter = num_to_col_letters(gr_last_col+1)
        gr_last_row = open_gr_sht.range(f'B'+ str(open_gr_sht.cells.last_cell.row)).end('up').row
        gr_date_col = gr_col_list.index("Date")+1
        gr_date_col_letter = num_to_col_letters(gr_date_col+1)
        
        open_gr_sht.api.AutoFilterMode=False
        # open_gr_sht.range(f"B{gr_head_row}:{gr_last_col_letter}{gr_last_row}").api.Sort(Key1=open_gr_sht.range(f"B{gr_head_row}:B{gr_last_row}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        open_gr_sht.range(f"B{gr_head_row}:{gr_last_col_letter}{gr_last_row}").api.Sort(Key1=open_gr_sht.range(f"B{gr_head_row}:B{gr_last_row}").api,
            Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        #Deleting Jrn Entries
        while not open_gr_sht.range(f"B{gr_head_row+1}").value.upper().startswith("MRN"):
            open_gr_sht.range(f"{gr_head_row+1}:{gr_head_row+1}").delete()

        mrn_last_row = open_gr_sht.range(f"{gr_date_col_letter}{gr_head_row}").end("down").row
        
        #Refreshing Source and data in Pivot
        
        wb.activate()
        pivot_sht.activate()
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"Sheet1!R{gr_head_row}C2:R{ mrn_last_row}C{gr_last_col}"#as list starts from B for for selecting  -1

        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()
        #Removing Gasoline from filter
        for product in wb.api.ActiveSheet.PivotTables(1).PivotFields("Product Name").PivotItems():
            if product.Name == "Ethanol":
                product.Visible = True
            else:
                product.Visible = False
        # try:
        #     wb.api.ActiveSheet.PivotTables(1).PivotFields("Product Name").PivotItems("Gasoline").Visible = False
        # except:
        #     pass
        #creating Sheet2 from Pivot Table
        try:
            wb.sheets("Sheet2").delete()
        except:
            pass
        pivot_last_row = pivot_sht.range("C4").end("down").row
        pivot_sht.range(f"C{pivot_last_row}").api.ShowDetail = True
        detail_sht = wb.app.selection.sheet.name
        mrn_df = wb.app.selection.expand("table").options(pd.DataFrame, 
                                header=1,
                                index=False
                                ).value
        #Copy pasting data in Sheet 4
        #getting required columns only 
        mrn_df = mrn_df[["Vendor Ref.", "Links", "Voucher", "Product Name", "Date", "BOLNumber", "Terminal ", "Account", "Billed Qty", "Credit Amount"]]
        sht_4 = wb.sheets("Sheet4")
        #Updating price formula
        mrn_df['price'] = mrn_df['Credit Amount']/mrn_df['Billed Qty']
        sht_4.api.AutoFilterMode=False
        sht_4.range("A2").expand("table").clear()
        sht_4.range("A2").options(pd.DataFrame, 
                                header=False,
                                index=False 
                                ).value = mrn_df 
        ##################Updating PIVOT Sheet Total Column#########################################
        wb.activate()
        pivot_sht.activate()
        pivot_sht_cols = pivot_sht.range("A4").expand("right").value     
        p_total_col_num = pivot_sht_cols.index("Total")+1#+1 for ignoring zero index
        p_total_col = num_to_col_letters(p_total_col_num)
        pivote_date_col_num = p_total_col_num - 1
        pivote_date_col = num_to_col_letters(pivote_date_col_num)
        p_row_label_num = pivot_sht_cols.index("Row Labels")+1
        p_row_label = num_to_col_letters(p_row_label_num)
        p_diff_num = pivot_sht_cols.index("Diff")+1
        p_diff_col = num_to_col_letters(p_diff_num)
        
        #Filling down and formating till lat row of pivot sheet
        pivot_sht.range(f"{pivote_date_col}5:{pivote_date_col}{pivot_last_row}").api.Select()
        wb.app.selection.api.FillDown()
        #Clearing extar data
        pivot_sht.range(f"{pivote_date_col}{pivot_last_row+1}:G{p_loc_r1 - 1}").clear()
        #adding sum formula
        pivot_sht.range(f"{pivote_date_col}{pivot_last_row}").formula = f"=SUM({pivote_date_col}5:{pivote_date_col}{pivot_last_row-1}"
        pivot_sht.range(f"{p_total_col}{pivot_last_row}").formula = f"=SUM({p_total_col}5:{p_total_col}{pivot_last_row-1}"
        ##Entering values based on location from dataframe
        location_list = pivot_sht.range("A5").expand("down").value
        for location in range(len(location_list)-1):
            if input_day <=15:
                pivot_sht.range(f"{p_total_col}{location+5}").value = sht3_dict[location_list[location]]
            else:
                pivot_sht.range(f"{p_total_col}{location+5}").value = f"={pivot_dict[location_list[location]]}*42"
            #Updating Diff Column
            pivot_sht.range(f"{p_diff_col}5:{p_diff_col}{len(location_list)+3}").select()#for selecting till last location row
            wb.app.selection.api.FillDown()
        #Sort Date Ascending order
        sht4_last_row = len(mrn_df)+2#{len(mrn_df)+2 +2 for zero index and excluding heading
        sht4_date_col = mrn_df.columns.get_loc("Date")+1
        sht4_date_letter = num_to_col_letters(sht4_date_col)
        sht4_term_col = mrn_df.columns.get_loc("Terminal ")+1
        sht4_term_letter = num_to_col_letters(sht4_term_col)
        sht_4.range(f"{p_row_label}1:K{sht4_last_row}").api.Sort(Key1=sht_4.range(f"{sht4_date_letter}1:{sht4_date_letter}{sht4_last_row}").api,
            Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)


        # #Updating mrns in Pivot2 sheet
        pivot2_sht = wb.sheets("Pivot2")
        pivot2_last_row = pivot2_sht.range(f"G{pivot2_sht.cells.last_cell.row}").end("up").row
        
        start_row = 2
        prev_row = 2
        i=2
        while i < pivot2_last_row:
            # if pivot2_sht.range(f"I{i}")!=0:
            #     start_row = i
            if "Total" in pivot2_sht.range(f"G{i}").value:
                start_row = prev_row
                
                if pivot2_sht.range(f"H{i}").value != 0:
                    #Filter terminal for inserting rows
                    sht4_term_col = mrn_df.columns.get_loc("Terminal ")+1
                    sht4_term_letter = num_to_col_letters(sht4_term_col)
                    
                    wb.activate()
                    sht_4.activate()
                    criteria_value = pivot2_sht.range(f"G{i-1}").value
                    sht_4.api.AutoFilterMode=False
                    sht_4.range(f"A1:K{sht4_last_row}").api.AutoFilter(Field:=f"{sht4_term_col}", Criteria1:=criteria_value)
                    if input_day <=15:
                        st_dt=input_datetime.replace(day=1)
                        
                    else:
                        st_dt=input_datetime.replace(day=16)
                    sht_4.range(f"A1:K{sht4_last_row}").api.AutoFilter(Field:=f"{sht4_date_col}", Criteria1:=[f'>={st_dt}'],
                                                                        Operator:=win32c.AutoFilterOperator.xlAnd, Criteria2:=[f'<={input_datetime}'])
                        
                    #Check if Filter contains row or not
                    if sht_4.range(f'A'+ str(sht_4.cells.last_cell.row)).end('up').row != 1:
                        sht_4.range(f"2:{sht4_last_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(sht_4.range(f"A{sht4_last_row+5}").api)
                    
                        sht_4.range(f"A{sht4_last_row+5}").expand('down').api.EntireRow.Copy()
                        wb.activate()
                        pivot2_sht.activate()
                        pivot2_sht.range(f"A{i}").api.EntireRow.Select()
                        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                        wb.app.api.CutCopyMode=False

                        #Deleting prev month rows
                        #getting range for deleting rows from range {prev_row}:{i-1}
                        delete_i = i-1
                        while int(input_month) == pivot2_sht.range(f"E{delete_i}").value.month or pivot2_sht.range(f"I{delete_i}").value != 0:
                            delete_i-=1
                            if pivot2_sht.range(f"E{delete_i}").value==None:
                                break
                            if delete_i == prev_row:
                                break
                        #Now checking for 0 entry
                        # while pivot2_sht.range(f"I{delete_i}").value != 0:
                        #     delete_i-=1
                        
                        if delete_i == i-1-1 or delete_i == prev_row:#Then no deletion
                            new_i = i + sht_4.range(f"A{sht4_last_row+5}").expand('down').rows.count
                        else:
                            #Delete
                            
                            delete_count = pivot2_sht.range(f"{prev_row}:{delete_i-1}").rows.count
                            pivot2_sht.range(f"{prev_row}:{delete_i-1}").api.Delete(Shift:=win32c.DeleteShiftDirection.xlShiftUp)

                            #Modifying i based on rows inserted
                            new_i = i + sht_4.range(f"A{sht4_last_row+5}").expand('down').rows.count - delete_count

                        #deleting copied data from sheet 4
                        pv_2_row_count = sht_4.range(f"A{sht4_last_row+5}").expand('down').rows.count
                        sht_4.range(f"A{sht4_last_row+5}").expand('down').api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                        
                        
                        #Updating total formulas
                        pivot2_sht.range(f"I{new_i}").formula = f"=+SUBTOTAL(9,I{start_row}:I{new_i-1})"
                        
                        pivot2_sht.range(f"M{new_i}").formula = f"=+SUBTOTAL(9,M{start_row}:M{new_i-1})"
                        pivot2_sht.range(f"N{new_i}").formula = f"=+SUBTOTAL(9,N{start_row}:N{new_i-1})"

                        #Updating otther columns
                        # pivot2_sht.range(f"M{i}:P{new_i}").formula = f"=+SUBTOTAL(9,N{start_row}:N{new_i-1})"

                        #Updating  Amount 	 Final Amount 	 Final Price 	 True-Up 	 Freight Rate 	 Freight Amount 
                        # pivot2_sht.range(f"M{start_row+1}:P{start_row+1}").copy(pivot2_sht.range(f"M{start_row+2}:P{new_i-1}"))
                        pivot2_sht.range(f"M{start_row+1}:P{start_row+1}").copy(pivot2_sht.range(f"M{new_i-pv_2_row_count}:P{new_i-1}"))
                        #updating O column Final Price forula
                        pivot2_sht.range(f"O{new_i-pv_2_row_count}").formula = f"=+K{new_i-pv_2_row_count}"
                        pivot2_sht.range(f"O{new_i-pv_2_row_count}").copy(pivot2_sht.range(f"O{new_i-pv_2_row_count}:O{new_i-1}"))
                        i=new_i

                        print("Selected")
                    if pivot2_sht.range(f"H{i}").value > 0: #Subtracting from existing BOLs
                        prev_row = i+1
                        #Ignoring row with zero value
                        while pivot2_sht.range(f"I{start_row}").value == 0:
                            start_row+=1
                        while pivot2_sht.range(f"I{start_row}").value <=  pivot2_sht.range(f"H{i}").value and start_row != i:
                            pivot2_sht.range(f"I{start_row}").formula = f'={pivot2_sht.range(f"I{start_row}").value} - {pivot2_sht.range(f"I{start_row}").value}'
                            pivot2_sht.range(f"I{start_row}").api.NumberFormat = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                            start_row += 1
                        if  pivot2_sht.range(f"I{start_row}").value >  pivot2_sht.range(f"H{i}").value:
                            pivot2_sht.range(f"I{start_row}").formula = f'={pivot2_sht.range(f"I{start_row}").value} - {pivot2_sht.range(f"H{i}").value}'
                            start_row = i+1
                        while pivot2_sht.range(f"H{i}").value !=0:
                            start_row-=1
                            start_value = float(pivot2_sht.range(f"I{start_row}").formula.split(' - ')[0].replace("=",""))
                            if start_value > pivot2_sht.range(f"H{i}").value:
                                pivot2_sht.range(f"I{start_row}").value = -pivot2_sht.range(f"H{i}").value
                            elif start_value < pivot2_sht.range(f"H{i}").value:
                                pivot2_sht.range(f"I{start_row}").value = -start_value
                    elif pivot2_sht.range(f"H{i}").value < 0:
                        while pivot2_sht.range(f"I{start_row}").value == 0:
                            start_row+=1
                        pivot2_sht.range(f"I{start_row}").formula = pivot2_sht.range(f"I{start_row}").formula+f'+{str(pivot2_sht.range(f"H{i}").value).replace("-","")}'
                        prev_row = i+1
                    else:
                        prev_row = i+1


                else:
                    prev_row=i+1
            i+=1
            pivot2_last_row = pivot2_sht.range(f"G{pivot2_sht.cells.last_cell.row}").end("up").row



        #Updating Litle Rock Tank Inventory(Little Rock Costing Tab)
        retry=0
        while retry < 10:
            try:
                lrti_wb = xw.Book(input_lrti_xl, update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        rack_costing_sht = lrti_wb.sheets["Rack Costing Working"]
        lr_costing_sht = wb.sheets["Little Rock Costing"]




        rack_costing_last_row = rack_costing_sht.range(f"A{rack_costing_sht.cells.last_cell.row}").end("up").row
        lr_costing_last_row = lr_costing_sht.range(f"A{lr_costing_sht.cells.last_cell.row}").end("up").row

        lr_cost_values = lr_costing_sht.range(f"A1:A{lr_costing_last_row}").value
        lr_atlas_1 = lr_cost_values.index('Atlas Date')+1#+1 for ignoring zero index
        lr_atlas_2 = lr_cost_values.index('Atlas Date', lr_atlas_1)+1#+1 for ignoring zero index


        li_col_list = lr_costing_sht.range(f"A2").expand("right").value
        li_atlas_date_col = li_col_list.index("Atlas Date")+1
        li_atlas_date_col_letter = num_to_col_letters(li_atlas_date_col)

        li_bol_col = li_col_list.index("BOL")+1
        li_bol_col_letter = num_to_col_letters(li_bol_col)

        li_qtgal_col = li_col_list.index("Qty.Gal")+1
        li_qtgal_col_letter = num_to_col_letters(li_qtgal_col)

        li_atqty_col = li_col_list.index("Atlas qty ")+1
        li_atqty_col_letter = num_to_col_letters(li_atqty_col)

        li_location_col = li_col_list.index("location")+1
        li_location_col_letter = num_to_col_letters(li_location_col)

        li_mrn_col = li_col_list.index("MRN#")+1
        li_mrn_col_letter = num_to_col_letters(li_mrn_col)

        li_deal_col = li_col_list.index("Deal#")+1
        li_deal_col_letter = num_to_col_letters(li_deal_col)

        li_date_col = li_col_list.index("Date")+1
        li_date_col_letter = num_to_col_letters(li_date_col)

        li_invdate_col = li_col_list.index("Inv Trf Date")+1
        li_invdate_col_letter = num_to_col_letters(li_invdate_col)

        li_vendor_col = li_col_list.index("Vendor")+1
        li_vendor_col_letter = num_to_col_letters(li_vendor_col)

        li_rate_col = li_col_list.index("Rate")+1
        li_rate_col_letter = num_to_col_letters(li_rate_col)

        li_fprice_col = li_col_list.index("Final Price")+1
        li_fprice_col_letter = num_to_col_letters(li_fprice_col)

        li_famount_col = li_col_list.index("Final Amount")+1
        li_famount_col_letter = num_to_col_letters(li_famount_col)

        # li_last_col_col = len(li_col_list)
        # li_last_col_letter = num_to_col_letters(li_last_col_col)
        li_last_col_letter = "AB" #Hard coded due to many gaps in column




        lr_filter_dict = {"BioUrja/North Magel":lr_atlas_1, "BioUrja/South Magel":lr_atlas_2}

        for key in lr_filter_dict.keys():
            rack_costing_last_row = rack_costing_sht.range(f"A{rack_costing_sht.cells.last_cell.row}").end("up").row
            lr_costing_last_row = lr_costing_sht.range(f"A{lr_costing_sht.cells.last_cell.row}").end("up").row

            lr_cost_values = lr_costing_sht.range(f"A1:A{lr_costing_last_row}").value
            lr_atlas_1 = lr_cost_values.index('Atlas Date')+1#+1 for ignoring zero index
            lr_atlas_2 = lr_cost_values.index('Atlas Date', lr_atlas_1)+1#+1 for ignoring zero index
            lr_filter_dict = {"BioUrja/North Magel":lr_atlas_1, "BioUrja/South Magel":lr_atlas_2}
            lrti_wb.activate()
            rack_costing_sht.activate()
            rack_costing_sht.api.AutoFilterMode=False
            rack_costing_last_row = rack_costing_sht.range(f"A{rack_costing_sht.cells.last_cell.row}").end("up").row
            if input_day <=15:
                st_dt=input_datetime.replace(day=1)
            else:
                st_dt=input_datetime.replace(day=16)
            rack_costing_sht.api.Range(f"A1:AD{rack_costing_last_row}").AutoFilter(Field:=1, Criteria1:=[f'>={st_dt}'],
                                                                        Operator:=win32c.AutoFilterOperator.xlAnd, Criteria2:=[f'<={input_datetime}'])
            # rack_costing_sht.api.Range(f"A1:AD{rack_costing_last_row}").AutoFilter(Field:=1, Criteria1:=[f'<={input_datetime}'])
            rack_costing_sht.api.Range(f"A1:AD{rack_costing_last_row}").AutoFilter(Field:=7, Criteria1:=key)
            
            if rack_costing_sht.range(f'A'+ str(rack_costing_sht.cells.last_cell.row)).end('up').row == 1:
                rack_costing_sht.api.AutoFilterMode=False
                if input_day <=15:
                    st_dt=input_datetime.replace(day=1)
                else:
                    st_dt=input_datetime.replace(day=16)
                rack_costing_sht.api.Range(f"A1:AD{rack_costing_last_row}").AutoFilter(Field:=1, Criteria1:=[f'>={st_dt}'],
                                                                            Operator:=win32c.AutoFilterOperator.xlAnd, Criteria2:=[f'<={input_datetime}'])
                # rack_costing_sht.api.Range(f"A1:AD{rack_costing_last_row}").AutoFilter(Field:=1, Criteria1:=[f'<={input_datetime}'])
                rack_costing_sht.api.Range(f"A1:AD{rack_costing_last_row}").AutoFilter(Field:=7, Criteria1:=key.replace(" ", ""))
            rack_costing_last_row = rack_costing_sht.range(f"A{rack_costing_sht.cells.last_cell.row}").end("up").row
            if rack_costing_sht.range(f'A'+ str(rack_costing_sht.cells.last_cell.row)).end('up').row != 1:
                rack_costing_sht.range(f"2:{rack_costing_last_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(lr_costing_sht.range(f"A{lr_costing_last_row+20}").api)


                data_st_row = lr_filter_dict[key]+2
                data_row = lr_costing_sht.range(f"A{data_st_row}").end("down").row
                if lr_costing_sht.range(f"A{data_row}").value =="South":
                    data_row = data_st_row
                elif data_st_row == lr_costing_last_row:
                    data_row = data_st_row
                lr_costing_sht.range(f"A{lr_costing_last_row+20}").expand('down').api.EntireRow.Copy()
                row_count = lr_costing_sht.range(f"A{lr_costing_last_row+20}").expand('down').size
                new_i = data_row + row_count
                wb.activate()
                lr_costing_sht.activate()
                lr_costing_sht.range(f"A{data_row+1}").api.EntireRow.Select()
                wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)
                lr_costing_sht.range(f"A{lr_costing_last_row+20+row_count}").expand('down').api.EntireRow.Delete(Shift:=win32c.DeleteShiftDirection.xlShiftUp)
                wb.app.api.CutCopyMode=False
                lr_costing_sht.range(f"O{data_row}:AB{data_row}").copy(lr_costing_sht.range(f"O{data_row+1}:AB{new_i}"))
                lr_costing_sht.range(f"O{data_row+1}:O{new_i}").value = None
                lr_costing_sht.range(f"R{data_row+1}:R{new_i}").value = None
                
                
                lr_costing_last_row = lr_costing_sht.range(f"A{lr_costing_sht.cells.last_cell.row}").end("up").row
                lr_cost_values = lr_costing_sht.range(f"A1:A{lr_costing_last_row}").value
                lr_atlas_1 = lr_cost_values.index('Atlas Date')+1#+1 for ignoring zero index
                lr_atlas_2 = lr_cost_values.index('Atlas Date', lr_atlas_1)+1#+1 for ignoring zero index
                lr_filter_dict = {"BioUrja/North Magel":lr_atlas_1, "BioUrja/South Magel":lr_atlas_2}

                #Sheet 4 logic for gettinglittlerocknorth terminal data
                wb.activate()
                sht_4.activate()
                sht_4.api.AutoFilterMode = False
                #Zeroing out Atals Quantity based on B1
                qty_match = "B1"
                # if key == "BioUrja/North Magel":
                qty_match = f"B{data_st_row-4}"
                if input_day <=15:
                    st_dt=input_datetime.replace(day=1)
                else:
                    st_dt=input_datetime.replace(day=16)
                sht_4.range(f"A1:K{sht4_last_row}").api.AutoFilter(Field:=f"{sht4_date_col}", Criteria1:=[f'>={st_dt}'],
                                                                        Operator:=win32c.AutoFilterOperator.xlAnd, Criteria2:=[f'<={input_datetime}'])
                if qty_match == "B0":
                    qty_match = "B1"
                    
                    #Applying filter in terminal
                    sht_4.range(f"A1:K{sht4_last_row}").api.AutoFilter(Field:=f"{sht4_term_col}", Criteria1:="NLITTLEROCKNORTH, AR")
                    sht4_last_row = sht_4.range(f'A'+ str(sht_4.cells.last_cell.row)).end('up').row
                else:
                    sht_4.range(f"A1:K{sht4_last_row}").api.AutoFilter(Field:=f"{sht4_term_col}", Criteria1:="NLITTLEROCKSOUTH, AR")
                    sht4_last_row = sht_4.range(f'A'+ str(sht_4.cells.last_cell.row)).end('up').row
                    
                if sht_4.range(f'A'+ str(sht_4.cells.last_cell.row)).end('up').row != 1:
                    sht_4.range(f"1:{sht4_last_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(lr_costing_sht.range(f"A{lr_costing_last_row+20}").api)
                    wb.activate()
                    lr_costing_sht.activate()
                    li_df = lr_costing_sht.range(f"A{lr_costing_last_row+20}:K{lr_costing_last_row+20}").expand("down").options(pd.DataFrame, 
                            header=1,
                            index=False 
                            ).value
                    lr_costing_sht.range(f"A{lr_costing_last_row+20}").expand("table").api.EntireRow.Delete(Shift:=win32c.DeleteShiftDirection.xlShiftUp)
                    row_count = len(li_df)
                    for i in range(row_count):
                        lr_costing_sht.range(f"A{new_i+1}").api.EntireRow.Select()
                        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)





                    #After row insertion add data
                    lr_costing_sht.range(f"{li_atlas_date_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Date"])
                    lr_costing_sht.range(f"{li_bol_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["BOLNumber"])
                    lr_costing_sht.range(f"{li_atqty_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Billed Qty"])
                    lr_costing_sht.range(f"{li_location_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Terminal "])
                    lr_costing_sht.range(f"{li_mrn_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Voucher"])
                    lr_costing_sht.range(f"{li_deal_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Links"])
                    lr_costing_sht.range(f"{li_date_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Date"])
                    lr_costing_sht.range(f"{li_invdate_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Date"])
                    lr_costing_sht.range(f"{li_vendor_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["Vend"])
                    lr_costing_sht.range(f"{li_rate_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["price"])
                    lr_costing_sht.range(f"{li_fprice_col_letter}{new_i+1}").options(transpose=True).value = list(li_df["price"])


                    #Updating formula
                    lr_costing_sht.range(f"{li_famount_col_letter}{new_i}:{li_last_col_letter}{new_i+row_count}").select()
                    wb.app.api.Selection.FillDown()

                    new_i += row_count


                try:
                    (lr_costing_sht.range(f"E{new_i+2}").value - lr_costing_sht.range(qty_match).value) >0
                except:
                    new_i += 1
                    while lr_costing_sht.range(f"E{new_i+2}").value is None:
                        new_i += 1
                if (lr_costing_sht.range(f"E{new_i+2}").value - lr_costing_sht.range(qty_match).value) >0: #Subtracting from existing BOLs
                    wb.activate()
                    lr_costing_sht.activate()
                # if (lr_costing_sht.range(f"E{new_i+3}").value - lr_costing_sht.range(qty_match).value) >0: #Subtracting from existing BOLs
                # if (lr_costing_sht.range(qty_match).value - lr_costing_sht.range(f"E{new_i+3}").value) >0: #Subtracting from existing BOLs
                # if (lr_costing_sht.range(f"E{new_i+3}").value !=0): #Subtracting from existing BOLs
                    #Logic for sorting data date wise
                    data_lst_row = lr_costing_sht.range(f"A{data_st_row}").end("down").row
                    lr_costing_sht.range(f"A{data_st_row}:AB{data_lst_row}").api.Sort(Key1=lr_costing_sht.range(f"A{data_st_row}:A{data_lst_row}").api,
                        Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
                    #Ignoring row with zero value
                    while lr_costing_sht.range(f"E{data_st_row}").value == 0:
                        data_st_row+=1
                    # while lr_costing_sht.range(f"E{data_st_row}").value <=   (lr_costing_sht.range(f"E{new_i+3}").value - lr_costing_sht.range(qty_match).value):
                    while lr_costing_sht.range(f"E{data_st_row}").value <=   lr_costing_sht.range(f"E{new_i+3}").value:
                        lr_costing_sht.range(f"E{data_st_row}").value = lr_costing_sht.range(f"E{data_st_row}").value - lr_costing_sht.range(f"E{data_st_row}").value
                        lr_costing_sht.range(f"E{data_st_row}").api.NumberFormat = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                        lr_costing_sht.range(f"D{data_st_row}").api.NumberFormat = 'General'
                        data_st_row += 1
                    if  lr_costing_sht.range(f"E{data_st_row}").value >  lr_costing_sht.range(f"E{new_i+3}").value:
                        lr_costing_sht.range(f"E{data_st_row}").value = lr_costing_sht.range(f"E{data_st_row}").value - lr_costing_sht.range(f"E{new_i+3}").value
                        data_st_row +=1

                    print("Done")
            else:
                # continue
                data_st_row = lr_filter_dict[key]+2
                data_row = lr_costing_sht.range(f"A{data_st_row}").end("down").row
                new_i = data_row
                qty_match = f"B{data_st_row-4}"
                if qty_match == "B0":
                    qty_match = "B1"
                try:
                    (lr_costing_sht.range(f"E{new_i+3}").value - lr_costing_sht.range(qty_match).value) >0
                except:
                    new_i += 1
                    while lr_costing_sht.range(f"E{new_i+2}").value is None:
                        new_i += 1
                if (lr_costing_sht.range(f"E{new_i+2}").value - lr_costing_sht.range(qty_match).value) >0: #Subtracting from existing BOLs
                    wb.activate()
                    lr_costing_sht.activate()
                # if (lr_costing_sht.range(f"E{new_i+3}").value - lr_costing_sht.range(qty_match).value) >0: #Subtracting from existing BOLs
                # if (lr_costing_sht.range(qty_match).value - lr_costing_sht.range(f"E{new_i+3}").value) >0: #Subtracting from existing BOLs
                # if (lr_costing_sht.range(f"E{new_i+3}").value !=0): #Subtracting from existing BOLs
                    #Logic for sorting data date wise
                    data_lst_row = lr_costing_sht.range(f"A{data_st_row}").end("down").row
                    lr_costing_sht.range(f"A{data_st_row}:AB{data_lst_row}").api.Sort(Key1=lr_costing_sht.range(f"A{data_st_row}:A{data_lst_row}").api,
                        Header =win32c.YesNoGuess.xlYes ,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
                    #Ignoring row with zero value
                    while lr_costing_sht.range(f"E{data_st_row}").value == 0:
                        data_st_row+=1
                    # while lr_costing_sht.range(f"E{data_st_row}").value <=   (lr_costing_sht.range(f"E{new_i+3}").value - lr_costing_sht.range(qty_match).value):
                    while lr_costing_sht.range(f"E{data_st_row}").value <=   lr_costing_sht.range(f"E{new_i+3}").value:
                        lr_costing_sht.range(f"E{data_st_row}").value = lr_costing_sht.range(f"E{data_st_row}").value - lr_costing_sht.range(f"E{data_st_row}").value
                        lr_costing_sht.range(f"E{data_st_row}").api.NumberFormat = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                        lr_costing_sht.range(f"D{data_st_row}").api.NumberFormat = 'General'
                        data_st_row += 1
                    if  lr_costing_sht.range(f"E{data_st_row}").value >  lr_costing_sht.range(f"E{new_i+3}").value:
                        lr_costing_sht.range(f"E{data_st_row}").value = lr_costing_sht.range(f"E{data_st_row}").value - lr_costing_sht.range(f"E{new_i+3}").value
                        data_st_row +=1

                    print("Done")

        lrti_wb.close()
        ##############Reconcilliation Part#################################################################################################
        wb.activate()
        open_gr_sht.activate()
        open_gr_cols = open_gr_sht.range(f"B6").expand("right").value
        open_gr_last_col_num = len(open_gr_cols)
        open_gr_voucher_col_num =  open_gr_cols.index("Voucher")
        open_gr_bol_col_num =  open_gr_cols.index("BOLNumber")
        open_gr_debit_col_num =  open_gr_cols.index("Debit Amount")
        open_gr_credit_col_num =  open_gr_cols.index("Credit Amount")
        open_gr_last_col = num_to_col_letters(open_gr_last_col_num+2)
        open_gr_voucher_col = num_to_col_letters(open_gr_voucher_col_num+2)
        open_gr_bol_col = num_to_col_letters(open_gr_bol_col_num+2)
        open_gr_debit_col = num_to_col_letters(open_gr_debit_col_num+2)
        open_gr_credit_col = num_to_col_letters(open_gr_credit_col_num+2)

        open_gr_last_row = open_gr_sht.range(f"{open_gr_voucher_col}{open_gr_sht.cells.last_cell.row}").end("up").row
        #Apply Duplicate filter on BOLNumber Column
        font_colour,Interior_colour = conditional_formatting_uniq(f"{open_gr_bol_col}:{open_gr_bol_col}",open_gr_sht,wb)
        open_gr_sht.api.AutoFilterMode=False
        open_gr_sht.api.Range(f"{open_gr_bol_col}6").AutoFilter(Field:=f"{open_gr_bol_col_num+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

        credit_range = open_gr_sht.range(f"{open_gr_credit_col}7").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible)
        debit_range = open_gr_sht.range(f"{open_gr_debit_col}7").end("down").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible)


        credit_count = credit_range.Count

        debit_count = debit_range.Count

        if debit_count != credit_count:
            open_gr_sht.api.AutoFilterMode=False
            #Filtering PVI for finding duplicate PVI
            open_gr_sht.range('G:G').api.FormatConditions.Delete()
            open_gr_sht.api.Range(f"{open_gr_voucher_col}6").AutoFilter(Field:=f"{open_gr_voucher_col_num+1}", Criteria1:="PVI*", Operator:=win32c.AutoFilterOperator.xlFilterValues)

            data_list = row_range_calc(open_gr_bol_col, open_gr_sht, wb)
            pvi_row = data_list[2][1].split(":")[0]

            # pvi_row= open_gr_sht.api.Range(f"{open_gr_voucher_col}7:{open_gr_voucher_col}{open_gr_last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address.split(':')[0].replace("$", "")
            font_colour,Interior_colour = conditional_formatting_uniq(f"{open_gr_bol_col}{pvi_row}:{open_gr_bol_col}{open_gr_last_row}", open_gr_sht, wb)
            open_gr_sht.api.Range(f"{open_gr_bol_col}6").AutoFilter(Field:=f"{open_gr_bol_col_num+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

        #Updating Difference in Accrual tab


        credit_range = open_gr_sht.range(f"{open_gr_credit_col}7").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible)
        debit_range = open_gr_sht.range(f"{open_gr_debit_col}7").end("down").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible)

        credit_sum = wb.api.Application.WorksheetFunction.Sum(credit_range)
        debit_sum = wb.api.Application.WorksheetFunction.Sum(debit_range)
        diff_amount = credit_sum - debit_sum

        open_gr_sht.api.AutoFilterMode=False

        accrual_sht = wb.sheets["Accrual"]


        wb.activate()
        accrual_sht.activate()
        accrual_last_row = accrual_sht.range(f"A{accrual_sht.cells.last_cell.row}").end("up").row
        accrual_last_row_2 = accrual_sht.range(f"X{accrual_sht.cells.last_cell.row}").end("up").row

        #updaint diffreence in accrual b column
        accrual_sht.range(f"B{accrual_last_row}").value = diff_amount

        #Deleting data in accrual sheet
        accrual_sht.range(f"A2:X{accrual_last_row_2}").delete()

        #Filtering out no fill values in open gr sheet and inserting them in accrual sheet
        wb.activate()
        open_gr_sht.activate()
        open_gr_sht.api.AutoFilterMode=False
        open_gr_sht.api.Range(f"{open_gr_bol_col}6").AutoFilter(Field:=f"{open_gr_bol_col_num+1}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterNoFill)
        # open_gr_last_row = open_gr_sht.range(f"{open_gr_voucher_col}{open_gr_sht.cells.last_cell.row}").end("up").row

        #Copy pasting filtered data in same sheet for insertion in accrual sheet
        open_gr_sht.range(f"{open_gr_voucher_col}7:{open_gr_last_col}7").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(open_gr_sht.range(f"A{open_gr_last_row+20}").api)

        #Copy pasting data in accrual sheet
        open_gr_sht.range(f"A{open_gr_last_row+20}:W{open_gr_last_row+20}").expand("down").api.EntireRow.Copy()
        wb.activate()
        accrual_sht.activate()
        accrual_sht.range(f"A2").api.EntireRow.Select()
        wb.app.api.Selection.Insert(Shift:=win32c.InsertShiftDirection.xlShiftDown)

        accr_last_mrn = accrual_sht.range(f"B2").end("down").row


        wb.api.ActiveSheet.PivotTables("PivotTable1").SourceData = f'Accrual!R1C1:R{accr_last_mrn}C23'
        wb.api.ActiveSheet.PivotTables("PivotTable1").PivotCache().Refresh()


        ###########################Getting open mrn as per date from pvi#####################################

        accrual_col_list = accrual_sht.range(f"A1").expand("right")
        accr_deliv_to_num = len(accrual_col_list)
        accr_deliv_to = num_to_col_letters(accr_deliv_to_num)

        accr_pvi_row = accr_last_mrn+2
        accr_pvi_date = accrual_sht.range(f"{accr_deliv_to}{accr_pvi_row}").value
        accr_pvi_year = accr_pvi_date.year
        accr_pvi_month_year = datetime.strftime(accr_pvi_date, "%m-%y")
        
        rack_date = datetime.strftime(last_day_of_month(accr_pvi_date.date()), "%Y%m%d")


        input_open_mrn = f"\\\\BIO-INDIA-FS\\India Sync$\\India\\{accr_pvi_year}\\{accr_pvi_month_year}\\Rack\\Open MRN Rack_{rack_date}.xlsx"#f'\\\\BIO-INDIA-FS\\India Sync$\\India\\{input_year}\\{input_month}-{input_year2}\\Little Rock Tank Inv Reco.xlsx'
        # input_open_mrn = curr_loc+r'\RackBackTrack\Raw Files'+f'\\OpenMRNRack{prev_mrn_date}.xlsx' #Open MRN Rack_20221130.xlsx
        if not os.path.exists(input_open_mrn):
            wb.save(output_location+f"\\RacbTrack_{input_date}.xlsx")
            return(f"{input_cta} Excel file not present, saving file till now with name RacbTrack_{input_date}.xlsx in putput folder")

        retry=0
        while retry < 10:
            try:
                open_mrn_wb = xw.Book(input_open_mrn, update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        open_mrn_sht = open_mrn_wb.sheets["Bills not received"]
        open_mrn_cols = open_mrn_sht.range(f"A1").expand("right").value
        open_mrn_del_to_num = open_mrn_cols.index("Delivery To")
        open_mrn_del_to = num_to_col_letters(open_mrn_del_to_num+1)
        
        sht_5 = wb.sheets["Sheet5"]
        sht_5.clear_contents()
        sht_5.clear_contents()
        open_mrn_sht.range(f"A1:{open_mrn_del_to}1").expand("down").copy(sht_5.range("A1"))
        open_mrn_wb.close()
        sht_5_last_row = sht_5.range(f"A{sht_5.cells.last_cell.row}").end("up").row
        # open_mrn_sht.copy(name="Sheet5", after=wb.sheets["Gasoline Costing"])
        #Copy paste pvi rows from accrual sheet
        accrual_sht.range(f"A{accr_pvi_row}:{accr_deliv_to}{accr_pvi_row}").expand("down").copy(sht_5.range(f"A{sht_5_last_row+1}"))

        sht_5_cols = sht_5.range(f"A1").expand("right").value
        sht_5_del_to_num = sht_5_cols.index("Delivery To")+1
        sht_5_del_to = num_to_col_letters(sht_5_del_to_num)
        sht_5_bol_col_num = sht_5_cols.index("BOLNumber")+1
        sht_5_bol_col = num_to_col_letters(sht_5_bol_col_num)
        sht_5_credit_col_num =  sht_5_cols.index("Credit Amount")+1
        sht_5_credit_col = num_to_col_letters(sht_5_credit_col_num)
        sht_5_debit_col_num =  sht_5_cols.index("Debit Amount")+1
        sht_5_debit_col = num_to_col_letters(sht_5_debit_col_num)


        #####ADD CONDTION FOR DUPLICATE PVI BOL$+###########################
        #Filtering duplicate bols
        sht_5.api.AutoFilterMode=False
        sht_5_last_row = sht_5.range(f"A{sht_5.cells.last_cell.row}").end("up").row
        font_colour,Interior_colour = conditional_formatting_uniq(f"{sht_5_bol_col}1:{sht_5_bol_col}{sht_5_last_row}", sht_5, wb)
        sht_5.api.Range(f"{sht_5_bol_col}1").AutoFilter(Field:=f"{sht_5_bol_col_num}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterCellColor)

        wb.activate()
        sht_5.activate()

        sht5_credit_range = sht_5.range(f"{sht_5_credit_col}7").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible)
        sht5_debit_range = sht_5.range(f"{sht_5_debit_col}7").end("down").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible)

        sht5_credit_sum = wb.api.Application.WorksheetFunction.Sum(sht5_credit_range)
        sht5_debit_sum = wb.api.Application.WorksheetFunction.Sum(sht5_debit_range)

        sht_5_diff_amount = sht5_credit_sum - sht5_debit_sum





        print("PVi data pasted")


        ######################Amking pivot of remaning mrn in sheet 5####################
        sht_5_last_row = sht_5.range(f"A{sht_5.cells.last_cell.row}").end("up").row
        sht_5.api.AutoFilterMode=False
        sht_5.api.Range(f"{sht_5_bol_col}1").AutoFilter(Field:=f"{sht_5_bol_col_num}", Criteria1:=Interior_colour, Operator:=win32c.AutoFilterOperator.xlFilterNoFill)

        #Copy pasting filtered data in same sheet for Pivot creation
        sht5_pivot_start = sht_5_last_row+10
        sht_5.range(f"A1:{sht_5_del_to}1").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy(sht_5.range(f"A{sht5_pivot_start}").api)
        sht5_pivot_end = sht_5.range(f"A{sht5_pivot_start}").end("down").row

        #Special case if pvi left and bol not matched then 
        pvi_amount_diff = 0 
        while sht_5.range(f"A{sht5_pivot_end}").value.startswith("PVI"):
            print("To be Handled")
            mrn_value = sht_5.range(f"U{sht5_pivot_end}").value
            pvi_debit = sht_5.range(f"{sht_5_debit_col}{sht5_pivot_end}").value
            sht_5.api.AutoFilterMode=False
            sht_5.api.Range(f"A{sht5_pivot_start}").AutoFilter(Field:=1, Criteria1:=mrn_value, Operator:=7)

            sht_5_last_row = sht_5.range(f"A{sht_5.cells.last_cell.row}").end("up").row

            mrn_credit = sht_5.range(f"{sht_5_credit_col}{sht_5_last_row}").value

            pvi_amount_diff += mrn_credit - pvi_debit

            #deleting mrn and pvi row from 2nd table
            sht_5.range(f"{sht_5_last_row}:{sht_5_last_row}").delete()
            sht_5.api.AutoFilterMode=False
            #Recalculating it as one row deleted
            sht5_pivot_end = sht_5.range(f"A{sht5_pivot_start}").end("down").row
            sht_5.range(f"{sht5_pivot_end}:{sht5_pivot_end}").delete()

            #reassigning sht5_pivot_end
            sht5_pivot_end = sht_5.range(f"A{sht5_pivot_start}").end("down").row







        #Updating diff in accrual sheet
        accrual_last_row = accrual_sht.range(f"A{accrual_sht.cells.last_cell.row}").end("up").row
        accrual_sht.range(f"B{accrual_last_row}").value = f"={diff_amount} + {sht_5_diff_amount} + {pvi_amount_diff}"


        #Creating Pivot for column Vendor Ref and Credit Amount
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Sheet5\'!R{sht5_pivot_start}C1:R{sht5_pivot_end}C{sht_5_del_to_num}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Sheet5'!R{sht5_pivot_end+5}C1", TableName="mrn_credit", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #logger.info("Adding particular Row Data in Pivot Table")
        PivotTable.PivotFields('Vendor Ref.').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Vendor Ref.').Position = 1

        #logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Credit Amount').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Credit Amount').NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        #Deleting Pivot from Accrual sheet Last month open MRN\
        accr_open_mrn_pivot_row = accrual_sht.range(f"A1").end("down").end("down").row+1
        accrual_sht.range(f"E{accr_open_mrn_pivot_row}").expand("table").clear_contents()


        sht_5.range(f"A{sht5_pivot_end+5}").expand('table').copy(accrual_sht.range(f"E{accr_open_mrn_pivot_row}"))

        pivot1_df = accrual_sht.range(f"A{accr_open_mrn_pivot_row}").expand("table").options(pd.DataFrame, 
                                header=1,
                                index=False 
                                ).value

        pivot2_df = accrual_sht.range(f"E{accr_open_mrn_pivot_row}").expand("table").options(pd.DataFrame, 
                                header=1,
                                index=False 
                                ).value
        #Remove last line containing Grand or Grnad Total
        pivot1_df = pivot1_df[:-1]
        pivot2_df = pivot2_df[:-1]

        # mrn_df = pd.merge(pivot1_df,pivot2_df, on="Row Labels", how="outer")

        df = pivot1_df.set_index('Row Labels').add(pivot2_df.set_index('Row Labels'), fill_value=0).reset_index()
        df.sort_values('Row Labels', inplace=True)

        combined_pivot_row = accrual_sht.range(f"E{accr_open_mrn_pivot_row}").end("down").end("down").row + 2

        pivot_color = accrual_sht.range(f"E{combined_pivot_row-1}").api.Interior.Color

        accrual_sht.range(f"E{combined_pivot_row}").expand("table").delete()
        accrual_sht.range(f"E{combined_pivot_row}").options(pd.DataFrame, 
                                header=None,
                                index=False 
                                ).value = df

        pivot_last_row = accrual_sht.range(f"E{accrual_sht.cells.last_cell.row}").end("up").row

        #Adding Total and it formula
        accrual_sht.range(f"E{pivot_last_row+1}").value = "TOTAL"
        accrual_sht.range(f"E{pivot_last_row+1}").api.Font.Bold = True
        accrual_sht.range(f"F{pivot_last_row+1}").formula = f"=SUM(F{combined_pivot_row}:F{pivot_last_row})"

        accrual_sht.range(f"E{pivot_last_row+1}:F{pivot_last_row+1}").api.Interior.Color = pivot_color



        ##################updating balance sheet amount############################
        bs_rack_inp = j_loc_bbr+f'\\BS Rack.xlsx'
        if not os.path.exists(bs_rack_inp):
            return(f"{bs_rack_inp} Excel file not present in raw files")

        bs_rack_df = pd.read_excel(bs_rack_inp, header=5, usecols=[1,2,3])

        bs_value = bs_rack_df.loc[bs_rack_df["Account Details"]=='    Open Goods Receipt']["Details - Curr. Period"].values[0]
        accrual_last_row = accrual_sht.range(f"A{accrual_sht.cells.last_cell.row}").end("up").row
        accrual_sht.range(f"B{accrual_last_row-2}").value = bs_value * -1


        #Updating total in B column
        pivot_1_last_row = accrual_sht.range(f"B{accr_open_mrn_pivot_row}").end("down").row

        accrual_total_row = accrual_sht.range(f"B{accrual_sht.cells.last_cell.row}").end("up").row

        accrual_sht.range(f"B{accrual_total_row}").formula = f"=SUM(B{pivot_1_last_row}:B{accrual_total_row-1})"


        ################Trueup Part##############################################################
        # if input_day == 15:
        print("Trueup Condition to be added")
        wb.activate()
        pivot2_sht.activate()

        pivot2_cols = pivot2_sht.range(f"A1").expand("right").value
        pivot2_last_col_num = len(pivot2_cols)
        pivot2_last_col = num_to_col_letters(pivot2_last_col_num)

        pivot2_po_col_num = pivot2_cols.index("Links")+1
        pivot2_po_col = num_to_col_letters(pivot2_po_col_num)

        pivot2_pdt_col_num = pivot2_cols.index("Pdt Name")+1
        pivot2_pdt_col = num_to_col_letters(pivot2_pdt_col_num)

        pivot2_bqty_col_num = pivot2_cols.index("Billed Qty")+1
        pivot2_bqty_col = num_to_col_letters(pivot2_bqty_col_num)

        pivot2_rate_col_num = pivot2_cols.index("Rate")+1
        pivot2_rate_col = num_to_col_letters(pivot2_rate_col_num)

        pivot2_amt_col_num = pivot2_cols.index("Amount")+1
        pivot2_amt_col = num_to_col_letters(pivot2_amt_col_num)
        
        pivot2_famt_col_num = pivot2_cols.index("Final Amount")+1
        pivot2_famt_col = num_to_col_letters(pivot2_famt_col_num)

        pivot2_fprice_col_num = pivot2_cols.index("Final Price")+1
        pivot2_fprice_col = num_to_col_letters(pivot2_fprice_col_num)

        pivot2_trueup_col_num = pivot2_cols.index("True-Up")+1
        pivot2_trueup_col = num_to_col_letters(pivot2_trueup_col_num)

        pivot2_date_col_num = pivot2_cols.index("Date")+1
        pivot2_date_col = num_to_col_letters(pivot2_date_col_num)

        pivot2_fbill_col_num = pivot2_cols.index("Freight Bill No.")+1
        pivot2_fbill_col = num_to_col_letters(pivot2_fbill_col_num)

        pivot2_frate_col_num = pivot2_cols.index("Freight Rate")+1
        pivot2_frate_col = num_to_col_letters(pivot2_frate_col_num)

        pivot2_fbdate_col_num = pivot2_cols.index("Freight Bill Booking Date")+1
        pivot2_fbdate_col = num_to_col_letters(pivot2_fbdate_col_num)

        pivot2_bol_col_num = pivot2_cols.index("BOLNumber")+1
        pivot2_bol_col = num_to_col_letters(pivot2_bol_col_num)




        #Applyling Filters
        pivot2_sht.api.AutoFilterMode=False
        
        pivot2_sht.range(f"A1:T{pivot2_last_row}").api.AutoFilter(Field:=pivot2_pdt_col_num, Criteria1:="Ethanol", Operator:=7)
        pivot2_sht.range(f"A1:T{pivot2_last_row}").api.AutoFilter(Field:=pivot2_bqty_col_num, Criteria1:='0', Operator:=7)

        #Clearing Contents from final amount to last column

        pivot2_last_row = pivot2_sht.range(f"G{pivot2_sht.cells.last_cell.row}").end("up").row
        if pivot2_last_row !=1:
            pivot2_sht.range(f"{pivot2_famt_col}2:{pivot2_last_col}{pivot2_last_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).ClearContents()

        #Now Filtering non zero billed qty
        pivot2_sht.range(f"A1:T{pivot2_last_row}").api.AutoFilter(Field:=pivot2_bqty_col_num, Criteria1:="<>0", Operator:=7)


        #Filling Formulas in 
        pivot2_last_row = pivot2_sht.range(f"G{pivot2_sht.cells.last_cell.row}").end("up").row

        pivot2_sht.api.Range(f"{pivot2_trueup_col}2:{pivot2_trueup_col}{pivot2_last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()

        #Apply current month Filter
        pivot2_last_row = pivot2_sht.range(f"G{pivot2_sht.cells.last_cell.row}").end("up").row
        pivot2_sht.range(f"A1:{pivot2_last_col}{pivot2_last_row}").api.AutoFilter(Field:=pivot2_date_col_num, 
        Criteria1:=f">={input_datetime.replace(day=1)}", Operator:=7)


        data = row_range_calc(pivot2_date_col, pivot2_sht, wb)
        first_row = data[0][0]
        #Updating Amount, Final Amount and Final Price and trueup
        pivot2_sht.range(f"{pivot2_amt_col}{first_row}").value = f"=+{pivot2_rate_col}{first_row}*{pivot2_bqty_col}{first_row}"
        pivot2_sht.range(f"{pivot2_famt_col}{first_row}").value = f"=+{pivot2_bqty_col}{first_row}*{pivot2_fprice_col}{first_row}"
        pivot2_sht.range(f"{pivot2_fprice_col}{first_row}").value = f"=+{pivot2_rate_col}{first_row}"
        pivot2_sht.range(f"{pivot2_trueup_col}{first_row}").value = f"=+{pivot2_famt_col}{first_row}-{pivot2_amt_col}{first_row}"


        #Filling dormulas in all filtered cells
        pivot2_sht.api.Range(f"{pivot2_amt_col}2:{pivot2_trueup_col}{pivot2_last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()

        print("Done")



        #Applying month filter for current

        pivot2_sht.api.AutoFilterMode=False
        
        ##########################Monthly Tureup part Check########################################################
        if input_date==last_date:#Montlhy True up condition
            #Create dataframe from both of this files and update trueup rate based on them rack_po_loc 
            
            # rack_po_dict = pd.read_excel(rack_po_loc).set_index("Vendor Name").to_dict()['PO#']
            TRUE_UP_DF = pd.read_excel(rack_po_loc)
            rack_po_dict = {}

            for i,x in TRUE_UP_DF.iterrows():

                rack_po_dict.setdefault(TRUE_UP_DF[TRUE_UP_DF.columns[0]][i], []).append(TRUE_UP_DF[TRUE_UP_DF.columns[1]][i])




            for key in rack_po_dict.keys():
                ap_t_df = pd.read_excel(truefile_loc, sheet_name = key,usecols="B:Q")
                
                if type(rack_po_dict[key]) is list:
                    value_list = [f"PO# {s}" for s in rack_po_dict[key]]
                    ap_price_dict = {}
                    ap_price_dict = ap_t_df[ap_t_df['Voucher'].isin(value_list)][['Voucher','Final Price']].set_index('Voucher').to_dict()['Final Price']
                else:
                    ap_price_dict = ap_t_df.loc[ap_t_df["Voucher"]==f"PO# {rack_po_dict[key]}"][['Voucher','Final Price']].set_index('Voucher').to_dict()['Final Price']
                wb.activate()
                pivot2_sht.activate()
                
                #Filtering po number for truep price update
                for po_key in ap_price_dict.keys():
                    pivot2_sht.api.AutoFilterMode=False
                    filt_key = po_key.replace("PO# ","POR:")
                    pivot2_sht.range(f"A1:{pivot2_last_col}{pivot2_last_row}").api.AutoFilter(Field:=pivot2_po_col_num, Criteria1:=filt_key, Operator:=7) #Links column containing po feild is B
                    #Updating final price for trueup calculation
                    flat_list, sp_lst_row,sp_address = row_range_calc(pivot2_fprice_col, pivot2_sht,wb)
                    
                    for address in sp_address:
                        address = pivot2_fprice_col + (f":{pivot2_fprice_col}").join(address.split(":"))

                        pivot2_sht.range(f"{address}").value = ap_price_dict[po_key]
            print("Done, updateed trueup prices")


        ###############################Freight update part check########################################################

        pivot2_sht.api.AutoFilterMode=False
        pivot2_sht.range(f"A1:{pivot2_last_col}{pivot2_last_row}").api.AutoFilter(Field:=pivot2_date_col_num, 
        Criteria1:=f">={input_datetime.replace(day=1)}", Operator:=7)
        # pivot2_sht.range(f"A1:{pivot2_last_col}{pivot2_last_row}").api.AutoFilter(Field:=pivot2_fbill_col_num,
        # Criteria1:=f"<>", Operator:=win32c.AutoFilterOperator.xlAnd, Criteria2:="<>No Freight")
        flat_list, sp_lst_row,sp_address = row_range_calc(pivot2_amt_col, pivot2_sht,wb)
        # sp_lst_row = pivot2_sht.range(f"{pivot2_amt_col}{pivot2_sht.cells.last_cell.row}").end("up").row
        # print("for loop for billed reading")
        # for i in flat_list:
        #     f_bill = pivot2_sht.range(f"{pivot2_fbill_col}{i}").value

        ####################

        frt_sht = wb.sheets("Frt GL")
        wb.activate()
        frt_sht.activate()
        frt_st_row = frt_sht.range("B1").end("down").end("down").row
        frt_lst_row = frt_sht.range(f"B{frt_sht.cells.last_cell.row}").end("up").row
        frt_col_list = frt_sht.range(f"B{frt_st_row}").expand('right').value
        frt_lst_col_num = len(frt_col_list)+1
        frt_lst_col = num_to_col_letters(frt_lst_col_num)
        frt_details_col_num = frt_col_list.index("Details")+2
        frt_details_col = num_to_col_letters(frt_details_col_num)
        frt_billno_col_num = frt_col_list.index("Bill No")+2
        frt_billno_col = num_to_col_letters(frt_billno_col_num)
        frt_qty_col_num = frt_col_list.index("Billed Qty")+2
        frt_qty_col = num_to_col_letters(frt_qty_col_num)
        frt_date_col_num = frt_col_list.index("Date")+2
        frt_date_col = num_to_col_letters(frt_date_col_num)
        frt_d_amount_col_num = frt_col_list.index("Debit Amount")+2
        frt_d_amount_col = num_to_col_letters(frt_d_amount_col_num)

        frt_new_col_num = len(frt_col_list)+2
        frt_new_col = num_to_col_letters(frt_new_col_num)

        frt_df = frt_sht.range(f"{frt_details_col}{frt_st_row}:{frt_d_amount_col}{frt_lst_row}").options(pd.DataFrame,
                                                                                                    header=1,
                                                                                                    index=False 
                                                                                                    ).value
        frt_df[['Bol No', 'Details']] = frt_df['Details'].str.split(';',expand=True)
        frt_df["Rate"] = frt_df["Debit Amount"] / frt_df["Billed Qty"]

        #Entering new columns
        frt_sht.range(f"{frt_new_col}{frt_st_row}").options(pd.DataFrame, 
                                header=1,
                                index=False 
                                ).value = frt_df[["Rate"]]

        frt_sht.range(f"A{frt_st_row}").options(pd.DataFrame, 
                                header=1,
                                index=False 
                                ).value = frt_df[["Bol No"]]

        #######Adding Vlookups ###################################
        #Bill No =VLOOKUP(F96,'Frt GL'!A:D,4,FALSE)
        #Date =VLOOKUP(F96, 'Frt GL'!A:E,5,FALSE)
        #Rate  =VLOOKUP(F96, 'Frt GL'!A:Z,26,FALSE)
        f_row = flat_list[0]
        pivot2_sht.range(f"{pivot2_fbill_col}{f_row}").formula = f'=IFERROR(VLOOKUP({pivot2_bol_col}{f_row},\'Frt GL\'!A:{frt_billno_col},{frt_billno_col_num},FALSE), "No Freight")'
        pivot2_sht.range(f"{pivot2_frate_col}{f_row}").formula = f'=IFERROR(VLOOKUP({pivot2_bol_col}{f_row},\'Frt GL\'!A:{frt_new_col},{frt_new_col_num},FALSE),0)'
        pivot2_sht.range(f"{pivot2_fbdate_col}{f_row}").formula = f'=IFERROR(VLOOKUP({pivot2_bol_col}{f_row},\'Frt GL\'!A:{frt_date_col},{frt_date_col_num},FALSE),"")'

        #Copy pasting formula
        wb.activate()
        pivot2_sht.activate()
        pivot2_sht.range(f"{pivot2_fbill_col}{f_row}:{pivot2_fbdate_col}{sp_lst_row}").select()
        wb.app.selection.api.FillDown()
        pivot2_sht.range(f"{pivot2_fbdate_col}{f_row}:{pivot2_fbdate_col}{sp_lst_row}").api.NumberFormat = "mm-dd-yyyy"

        pivot2_sht.range(f"{pivot2_frate_col}{f_row}:{pivot2_frate_col}{sp_lst_row}").select()
        wb.app.selection.api.FillDown()

        #Copy pasting data to replace formulas
        for add_range in sp_address:
            rate_address = pivot2_frate_col + (f":{pivot2_frate_col}").join(add_range.split(":"))
            bill_d_address = pivot2_fbill_col + (f":{pivot2_fbdate_col}").join(add_range.split(":"))
            pivot2_sht.range(f"{rate_address}").api.Copy()
            pivot2_sht.range(f"{rate_address}").api._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
            pivot2_sht.range(f"{rate_address}").api.NumberFormat = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

            pivot2_sht.range(f"{bill_d_address}").api.Copy()
            pivot2_sht.range(f"{bill_d_address}").api._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)



        #Removing filters from all sheets
        open_gr_sht.api.AutoFilterMode=False
        pivot2_sht.api.AutoFilterMode=False
        sht_5.api.AutoFilterMode=False
        accrual_sht.api.AutoFilterMode=False
        sht_4.api.AutoFilterMode=False
        sht6.api.AutoFilterMode=False


        wb.save(output_location+f"\\RackBackTrack_{input_date}.xlsx")
        print("Done")
        return f"RackBack Track file for date: {input_date} has been generated"
    except Exception as e:
        print(e)
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

# msg = rackbacktrack('01.31.2023', '01.31.2023')
# print(msg)