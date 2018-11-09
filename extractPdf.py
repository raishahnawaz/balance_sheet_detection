import pdftables_api
#import tabula
import xlrd
import pandas as pd
import numpy
import glob
import os


def CalculateDE(df):
    '''CODE TO ADD //Check for currency signs in the table (i.e $ in our case)'''
    Currency = "USD"
    tuplesList=list()
    IndexList=[29]
    TempStr="{"
    for x in IndexList:
        temp = df.iloc[x, :]
        print("LongtermDebt:", temp)
        # a = [numpy.NAN if x.value == '' else x.value for x in LongtermDebt]
        a = [x.value for x in temp if x.value != '']
        #print("LongtermDebt Isnull", a)
        TempStr+="\""
        TempStr+=a[0]+"\":"+str(a)+","

    TempStr =TempStr[0:-1]+"}"

    print("TempStr:" ,TempStr)
    print ("tuplesList",tuplesList)
    ResultsDf = pd.DataFrame(TempStr)
    print("Data Frame", ResultsDf.head(2))

    '''CODE TO ADD //Look for units on Balance Sheet table (i.e thousands, millions)'''
    Denomination = df.iloc[2, 0]
    print("Denomination:", Denomination)
    months = df.iloc[3, :].dropna()
    print("Months:", months,type(months))
    years = df.iloc[4, :]
    print("\nYears:", years)

    #LongtermDebt = df[df['one'] == 'Long-term debt']
    #print("LongtermDebt:", LongtermDebt)



    print(ResultsDf.head(2))

    ShortTermDebt = df.iloc[31, :]
    print("ShortTermDebt:", ShortTermDebt)

    ShareholderEquity = df.iloc[39, :]
    print("ShareholderEquity:", ShareholderEquity)

    ShareholderLoan = None
    Outputlabels = ["Date of Financials used", "Debt", "Long term borrowings", "Short term borrowings", "Equity",
                    "Shareholder's Equity", "Shareholder's loan", \
                    "Final Debt", "Final Equity", "DE ratio"]
    row=[]

    years_average_DE_ratio = None


path=r"C:\Users\Rai Shahnawaz\Desktop\AI Challenge\reports"
my_pdftables_api_key="cvmm56lxabx7"

if __name__ == '__main__':

    '''READ ALL PDFs FROM A DIRECTORY AND CONVERT ALL FILES INTO XLXS FORMAT USING pdftables-api'''

    pdfFiles = []
    check = input('Please enter the folder path where the files are located: ')
    src=path
    if not os.path.isdir(src):
        print('Invalid given path.')
        exit(1)

    path = os.path.join(src,"*.pdf")
    for file in glob.glob(path):
        pdfFiles.append(file)

    print("PDF Files:", pdfFiles)

    for file in pdfFiles:
        print (file, "   ")
        c = pdftables_api.Client(my_pdftables_api_key)
        #c.xlsx(file, file[0:len(file)-4]+'.xlsx')


    '''OPEN A CONVETED REPORT AND LOOK FOR THE BALANCE SHEET'''
    xl_workbook = xlrd.open_workbook('output.xlsx')
    Names=xl_workbook.sheet_names()
    for sheetname in Names:
         if sheetname=="Page 5":
             BalanceSheet=xl_workbook.sheet_by_name(sheetname)

    PageRows=list()
    for rowN in range(BalanceSheet.nrows):
         PageRows.append(BalanceSheet.row(rowN))

    labels=['one','two','three','four','five']
    df = pd.DataFrame.from_records(PageRows,columns=labels)

    '''PASS BALANCE SHEET FOR PARSING AND OUTPUT'''
    OutputDF=CalculateDE(df)


            # Read pdf into DataFrame
    #tabula.convert_into("testCopy.pdf", "output.csv", output_format="csv")

    #df = tabula.read_pdf("testCopy.pdf")
    #tabula.show()