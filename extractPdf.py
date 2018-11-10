import pdftables_api
#import tabula
import xlrd
import pandas as pd
import numpy
import glob
import os


path=r"C:\Users\Rai Shahnawaz\Desktop\AI Challenge\reports"
my_pdftables_api_key="cvmm56lxabx7"

def convert_pdfFiles_to_xlsx(src):
    pdfFiles = []
    path = os.path.join(src,"*.pdf")
    for file in glob.glob(path):
        pdfFiles.append(file)

    print("PDF Files:", pdfFiles)
    for file in pdfFiles:
        c = pdftables_api.Client(my_pdftables_api_key)
        #c.xlsx(file, file[0:len(file)-4]+'.xlsx')

def load_balance_sheet_toDF(filename):
    xl_workbook = xlrd.open_workbook(filename)
    Names=xl_workbook.sheet_names()
    for sheetname in Names:
         if sheetname=="Page 5":
             BalanceSheet=xl_workbook.sheet_by_name(sheetname)

    PageRows=list()
    for rowN in range(BalanceSheet.nrows):
         PageRows.append(BalanceSheet.row(rowN))

    labels=['one','two','three','four','five']
    return pd.DataFrame.from_records(PageRows,columns=labels)

def CalculateDE(DE_output_DF):
    print("Not implemented yet")
    ShareholderLoan = None
    Outputlabels = ["Date of Financials used", "Debt", "Long term borrowings", "Short term borrowings", "Equity",
                    "Shareholder's Equity", "Shareholder's loan", \
                    "Final Debt", "Final Equity", "DE ratio"]

    years_average_DE_ratio = None

def parse_balance_sheet_for_DE(df):
    '''CODE TO ADD //Check for currency signs in the table (i.e $ in our case)'''
    Currency = "USD"
    tuplesList=dict()
    IndexList=[29,31,39]
    for x in IndexList:
        temp = df.iloc[x, :]
        print("LongtermDebt:", temp)
        # a = [numpy.NAN if x.value == '' else x.value for x in LongtermDebt]
        a = [x.value for x in temp if x.value != '']
        tuplesList[a[0]]=a[1:]

    months = df.iloc[3, :]
    a = [x.value for x in months if x.value != '']
    tuplesList["months"] = a
    years = df.iloc[4, :]
    a = [x.value for x in years if x.value != '']
    tuplesList["years"] = a

    ResultsDf = pd.DataFrame(tuplesList)
    print("Data Frame", ResultsDf.head(2))

    '''CODE TO ADD //Look for units on Balance Sheet table (i.e thousands, millions)'''
    Denomination = df.iloc[2, 0]
    print("Denomination:", Denomination)

    writer = pd.ExcelWriter('outputFile.xlsx')
    ResultsDf.to_excel(writer, sheet_name='Results',header=True)
    writer.save()
    return ResultsDf

if __name__ == '__main__':

    '''READ ALL PDFs FROM A DIRECTORY AND CONVERT ALL FILES INTO XLXS FORMAT USING pdftables-api'''

    check = input('Please enter the folder path where the files are located: ')
    src=path
    if not os.path.isdir(src):
        print('Invalid given path.')
        exit(1)

    convert_pdfFiles_to_xlsx(src)

    '''OPEN A CONVERTED REPORT AND LOOK FOR THE BALANCE SHEET'''
    filename='output.xlsx'
    Balance_sheet_DF=load_balance_sheet_toDF(filename)

    '''BALANCE SHEET PARSING FOR DEBT EQUITY VALUES'''
    DE_output_DF=parse_balance_sheet_for_DE(Balance_sheet_DF)

    '''Calculating all the required fields using debt equity values'''
    OutputDF=CalculateDE(DE_output_DF)

