from codecs import lookup
from tabula import read_pdf
import PyPDF2
import pandas as pd
import xlwings as xw
import os, time
from datetime import datetime, timedelta
import re
import tabula as tb
import xlsxwriter
from tabula import read_pdf
from openpyxl import load_workbook
import openpyxl as xl
from datetime import datetime

# date_df = read_pdf(loc, pages = 1, guess = False, stream = True ,

#                                     pandas_options={'header':0}, area = ["30,290,120,415"], columns=["320"])[0]

# file1 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-BOFA Bulk.pdf","rb")
# file2 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-BOFA Rack.pdf","rb")
# file3 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-RCP South BOFA.pdf","rb")
# file4 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-RCP Midwest.pdf","rb")
# file5 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-BOFA-BU Export.pdf","rb")
# file6 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-BOFA-BioUrja Nehme Commodities.pdf","rb")
# file7 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-BOFA-BioUrja Energy Commodities.pdf","rb")
# file8 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-RCP Holdings.pdf","rb")
# file9 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-RCP South Chase.pdf","rb")
# file10 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-CHASE Bulk.pdf","rb")
# file11 = open("C:\\DEEPFOLDER\\Tasks\\BBR_PROCESS\\BBR_20221130\\Bank\\20221130-CHASE Rack.pdf","rb")

def BOFA_Bulk(wb,ws,file1):
 try:
     pdfReader = PyPDF2.PdfFileReader(file1)
     print(pdfReader.numPages)
     page = pdfReader.getPage(0)
     pdfData = page.extractText()
     print(pdfData)
     search_term = "Register Balance as of"

     if search_term in pdfData:
        print("found")
        line = pdfData[pdfData.find(search_term):].split('\n')[0]
     #    re.split('\s+',line)
        print(line)
        w=line.rstrip().split(" ")[-1]
        print(w)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    num1 = num[2:]
     #    print(num1)
     #    listtostr = ''.join(map(str,num1))
     #    print(listtostr)

     #    w =re.split('\s+',line)
     #    start = line.index('Register Balance as of')
     #    end = line.index('0')
     #    substring = line[start+30:end]
     #    sub =float(substring.strip().replace(" ",""))
     #    print(sub)
     #    wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
     #    ws = wb.sheets("Cash")
        ws.range('G10').value = w

 except Exception as e:
     raise e


def BOFA_Rack(wb,ws,file2):
     try:
      pdfReader = PyPDF2.PdfFileReader(file2)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      search_term = "Reconciled Transaction Total"

      if search_term in pdfData:
        print("found")
        line = pdfData[pdfData.find(search_term):].split('\n')[0]
        # re.split('\s+',line)
        print(line)
        w=line.rstrip().split(" ")[-4].replace(":","")
        print(w)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
        wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
        ws = wb.sheets("Cash")
        ws.range('G11').value = w

     except Exception as e:
          raise e

def BOFA_RCP_SOUTH(wb,ws,file3):

     try:
      pdfReader = PyPDF2.PdfFileReader(file3)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      date_df = read_pdf(file3, pages = 1, guess = False, stream = True ,

                                     pandas_options={'header':0}, area = ["13,198,370,385"], columns=["360"])[0]
      
      print(date_df)
      q=date_df.iloc[-1][0].replace("$","").replace(",","")
      print(q)
     #  search_term = "Adjusted Account Balance"

     #  if search_term in pdfData:
     #    print("found")
     #    line = pdfData[pdfData.find(search_term):].split('\n')[0]
     #    print(line)
        # re.split('\s+',line)
     #    print(line)
     #    w=line.rstrip().split(" ")[-1]
     #    print(w)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)

      wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
      ws = wb.sheets("Cash")
      ws.range('G12').value = q

     except Exception as e:
          raise e

def BOFA_RCP_MIDWEST(wb,ws,file4):

     try:
      pdfReader = PyPDF2.PdfFileReader(file4)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      date_df = read_pdf(file4, pages = 1, guess = False, stream = True ,

                                     pandas_options={'header':0}, area = ["13,198,370,385"], columns=["360"])[0]
      
      print(date_df)
      q=date_df.iloc[-1][0].replace("$","").replace(",","")
      print(q)
     #  search_term = "Adjusted Account Balance"

     #  if search_term in pdfData:
     #    print("found")
     #    line = pdfData[pdfData.find(search_term):].split('\n')[0]
     #    # re.split('\s+',line)
     #    print(line)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
      wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
      ws = wb.sheets("Cash")
      ws.range('G13').value = q

     except Exception as e:
          raise e

def BOFA_BU_EXPORT(wb,ws,file5):

     try:
      pdfReader = PyPDF2.PdfFileReader(file5)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      search_term = "Ending Balance"

      if search_term in pdfData:
        print("found")
        line = pdfData[pdfData.find(search_term):].split('\n')[0]
        # re.split('\s+',line)
        print(line)
        num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
        print(num)
        listtostr = ''.join(map(str,num))
        print(listtostr)
        wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
        ws = wb.sheets("Cash")
        ws.range('G14').value = listtostr

     except Exception as e:
          raise e

def BOFA_BU_NEHMA(wb,ws,file6):

     try:
      pdfReader = PyPDF2.PdfFileReader(file6)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      date_df = read_pdf(file6, pages = 1, guess = False, stream = True ,

                                     pandas_options={'header':0}, area = ["13,198,370,385"], columns=["360"])[0]
      
      print(date_df)
      q=date_df.iloc[-1][0].replace("$","").replace(",","")
      print(q)
     #  search_term = "Adjusted Account Balance"

     #  if search_term in pdfData:
     #    print("found")
     #    line = pdfData[pdfData.find(search_term):].split('\n')[0]
     #    # re.split('\s+',line)
     #    print(line)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
      wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
      ws = wb.sheets("Cash")
      ws.range('G15').value = q

     except Exception as e:
          raise e

def BOFA_BU_ENERGY(wb,ws,file7):

     try:
      pdfReader = PyPDF2.PdfFileReader(file7)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      date_df = read_pdf(file7, pages = 1, guess = False, stream = True ,

                                     pandas_options={'header':0}, area = ["13,198,370,385"], columns=["360"])[0]
      
      print(date_df)
      q=date_df.iloc[-1][0].replace("$","").replace(",","")
      print(q)
     #  search_term = "Adjusted Account Balance"

     #  if search_term in pdfData:
     #    print("found")
     #    line = pdfData[pdfData.find(search_term):].split('\n')[0]
     #    # re.split('\s+',line)
     #    print(line)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
      wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
      ws = wb.sheets("Cash")
      ws.range('G9').value = q

     except Exception as e:
          raise e

def RCP_HOLDINGS(wb,ws,file8):

     try:
      pdfReader = PyPDF2.PdfFileReader(file8)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      date_df = read_pdf(file8, pages = 1, guess = False, stream = True ,

                                     pandas_options={'header':0}, area = ["13,198,370,385"], columns=["360"])[0]
      
      print(date_df)
      q=date_df.iloc[-1][0].replace("$","").replace(",","")
      print(q)
     #  search_term = "Adjusted Account Balance"

     #  if search_term in pdfData:
     #    print("found")
     #    line = pdfData[pdfData.find(search_term):].split('\n')[0]
     #    # re.split('\s+',line)
     #    print(line)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
     #  wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
     #  ws = wb.sheets("Cash")
      ws.range('G17').value = q

     except Exception as e:
          raise e

def RCP_South_Chase(wb,ws,file9):

     try:
      pdfReader = PyPDF2.PdfFileReader(file9)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      date_df = read_pdf(file9, pages = 1, guess = False, stream = True ,

                                     pandas_options={'header':0}, area = ["13,198,370,385"], columns=["360"])[0]
      
      print(date_df)
      q=date_df.iloc[-1][0].replace("$","").replace(",","")
      print(q)
     #  search_term = "Adjusted Account Balance"

     #  if search_term in pdfData:
     #    print("found")
     #    line = pdfData[pdfData.find(search_term):].split('\n')[0]
     #    # re.split('\s+',line)
     #    print(line)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
     #  wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
     #  ws = wb.sheets("Cash")
      ws.range('G8').value = q

     except Exception as e:
          raise e

def Chase_Bulk(wb,ws,file10):

     try:
      pdfReader = PyPDF2.PdfFileReader(file10)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      search_term = "Register Balance as on"

      if search_term in pdfData:
        print("found")
        line = pdfData[pdfData.find(search_term):].split('\n')[0]
        # re.split('\s+',line)
        print(line)
        w=line.rstrip().split(" ")[-1]
        print(w)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
     #    wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
     #    ws = wb.sheets("Cash")
        ws.range('G6').value = w

     except Exception as e:
          raise e


def Chase_Rack(wb,ws,file11):

     try:
      pdfReader = PyPDF2.PdfFileReader(file11)
      print(pdfReader.numPages)
      page = pdfReader.getPage(0)
      pdfData = page.extractText()
      print(pdfData)
      search_term = "Register Balance as on"

      if search_term in pdfData:
        print("found")
        line = pdfData[pdfData.find(search_term):].split('\n')[0]
        # re.split('\s+',line)
        print(line)
        w=line.rstrip().split(" ")[-1]
        print(w)
     #    num = re.findall(r'[-+]?(?:\d*\.\d+|\d+)',line)
     #    print(num)
     #    listtostr = ''.join(map(str,num))
     #    print(listtostr)
     #    wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
     #    ws = wb.sheets("Cash")
        ws.range('G7').value = w

     except Exception as e:
          raise e

# def main():
#  BOFA_Bulk()
#  BOFA_Rack()
#  BOFA_RCP_SOUTH()
#  BOFA_RCP_MIDWEST()
#  BOFA_BU_EXPORT()
#  BOFA_BU_NEHMA()
#  BOFA_BU_ENERGY()
#  RCP_HOLDINGS()
#  RCP_South_Chase()
#  Chase_Bulk()
#  Chase_Rack()
#  main()

# if __name__ == "__main__":
#     try:
#       BOFA_Bulk()
#       BOFA_Rack()
#       BOFA_RCP_SOUTH()
#       BOFA_RCP_MIDWEST()
#       BOFA_BU_EXPORT()
#       BOFA_BU_NEHMA()
#       BOFA_BU_ENERGY()
#       RCP_HOLDINGS()
#       RCP_South_Chase()
#       Chase_Bulk()
#       Chase_Rack()
 
#     except Exception as e:
#         raise e


def cash(start_date,end_date):
     try:
        start_date2 = datetime.strftime(datetime.strptime(start_date2,"%m.%d.%Y"), "%Y%m%d")
     #    start_date1 = datetime.strftime(datetime.strptime(start_date,"%m.%d.%Y"), "%Y.%m.%d")
     #    end_date2 = datetime.strftime(datetime.strptime(end_date,"%m.%d.%Y"), "%Y-%m-%d")
     #    adte = date.today()
        file1 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-BOFA Bulk.pdf","rb")
        file2 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-BOFA Rack.pdf","rb")
        file3 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-RCP South BOFA.pdf","rb")
        file4 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-RCP Midwest.pdf","rb")
        file5 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-BOFA-BU Export.pdf","rb")
        file6 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-BOFA-BioUrja Nehme Commodities.pdf","rb")
        file7 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-BOFA-BioUrja Energy Commodities.pdf","rb")
        file8 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-RCP Holdings.pdf","rb")
        file9 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-RCP South Chase.pdf","rb")
        file10 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-CHASE Bulk.pdf","rb")
        file11 = open(f"J:\India\BBR\2023\BBR_{start_date2}\Bank\{start_date2}-CHASE Rack.pdf","rb")
        wb = xw.Book(f"J:\India\BBR\2023\BBR_{start_date2}\BioUrja - Consolidated Borrowing Base Syndication {start_date}_Working.xlsx")
        retry = 0
        while retry < 10:
            try:
                wb = xw.Book("C:\DEEPFOLDER\Tasks\BBR_PROCESS\BioUrja - Consolidated Borrowing Base Syndication 11.30.2022_Working.xlsx")
                break            
            except Exception as e:
                time.sleep(5)
                retry+=1                
                if retry ==10:
                    raise e
        ws = wb.sheets("Cash")
        BOFA_Bulk(wb,ws,file1)
        BOFA_Rack(wb,ws,file2)
        BOFA_RCP_SOUTH(wb,ws,file3)
        BOFA_RCP_MIDWEST(wb,ws,file4)
        BOFA_BU_EXPORT(wb,ws,file5)
        BOFA_BU_NEHMA(wb,ws,file6)
        BOFA_BU_ENERGY(wb,ws,file7)
        RCP_HOLDINGS(wb,ws,file8)
        RCP_South_Chase(wb,ws,file9)
        Chase_Bulk(wb,ws,file10)
        Chase_Rack(wb,ws,file11)

        return(f"Cash report for {start_date} has been generated successfully")
        


     except Exception as e:
          raise e


cash('01.15.2023','01.15.2023')

