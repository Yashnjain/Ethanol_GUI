import os
import re
import PyPDF2
import xlwings as xw
from codecs import lookup
from tabula import read_pdf
from datetime import datetime


# file = open("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BBR_20221130\Future - BNP.pdf","rb")
# file2 = open("C:\DEEPFOLDER\Tasks\BACKUP\BBR_BACKUP\BBR_20221130\Future - Macquarie.pdf","rb")


def BNP(file,wb,ws):
    try:
     pdfReader = PyPDF2.PdfFileReader(file)
     print(pdfReader.numPages)
     page = pdfReader.getPage(0)
     pdfData = page.extractText()
     print(pdfData)
     search_term = "NET LIQUIDATING VALUE"

     if search_term in pdfData:
        print("found")
        line = pdfData[pdfData.find(search_term):].split('\n')[0]
        print(line)
        we =re.split('\s+',line)
        start = line.index('NET LIQUIDATING VALUE')
        end = line.index('CR')
        substring = line[start+21:end]
        sub =float(substring.strip().replace(" ",""))
        print(sub)
        # wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
        # ws = wb.sheets("NLV Futures")
        ws.range('G7').value = sub

        # num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
        # print(num)
        # num1 = num[:3]
        # print(num1)
        # listtostr = ''.join(map(str,num1))
        # print(listtostr)

     else:
        quit

    except Exception as e:
     raise e


def macquarie(file2,wb,ws):
    try:
     pdfReader = PyPDF2.PdfFileReader(file2)
     print(pdfReader.numPages)
     page = pdfReader.getPage(0)
     pdfData = page.extractText()
     print(pdfData)
     search_term = "Net Liquidating Value"

    #  req_data_re = re.compile(r'^\d{5} [A-Z].*')
    #  for line in pdfData.split('/n'):
    #     if req_data_re.match(line):
    #         print(line)

     if search_term in pdfData:
        print("found")
        line = pdfData[pdfData.find(search_term):].split('\n')[0]
        # re.split('\s+',line)
        print(line)
        num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
        print(num)
        listtostr = ''.join(map(str,num))
        print(listtostr)
        # wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
        # ws = wb.sheets("NLV Futures")
        ws.range('G6').value = listtostr

        # acc_no = a[a.find('Account'):a.find('Account')+17]
        # for line in file:
            # if re.findall(search_term,line) in line:
            #     print(line)

     else:
       quit

    


    except Exception as e:
     raise e


def NLV_FUTURESSS(start_date,end_date):
    try:
        start_date2 = datetime.strftime(datetime.strptime(start_date,"%m.%d.%Y"), "%Y%m%d")
        
        # end_date2 = datetime.strftime(datetime.strptime(end_date,"%d.%m.%Y"), "%Y%m%d")
        future_bnp = r"J:\\India\\BBR\\IT_BBR\\Reports\\NLV\\INPUT\\Future - BNP.pdf"
        future_macquarie = r"J:\\India\\BBR\\IT_BBR\\Reports\\NLV\\INPUT\\Future - Macquarie.pdf" 
        wb_file = f"J:\\India\\BBR\\IT_BBR\\Reports\\NLV\\INPUT\\BioUrja - Consolidated Borrowing Base Syndication {start_date}.xlsx"
        file = open(future_bnp,"rb")
        if not os.path.exists(future_bnp):
            return(f"{file} Input Excel file not present")
        file2 = open(future_macquarie,"rb")
        if not os.path.exists(future_macquarie):
            return(f"{file2} Input Excel file not present")
        wb = xw.Book( wb_file)
        if not os.path.exists( wb_file):
            return(f"{ wb_file} Input Excel file not present")
        ws = wb.sheets("NLV Futures")
        macquarie(file2,wb,ws)
        BNP(file,wb,ws)
        return(f"NLV Futures report for {start_date} has been generated successfully")

    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass
# NLV_FUTURES('15.01.2023','15.01.2023')

