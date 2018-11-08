import pdftables_api
import tabula
import xlrd
import pandas as pd


def CalculateDE(df):
    '''CODE TO ADD //Check for currency signs in the table (i.e $ in our case)'''
    Currency = "USD"

    '''CODE TO ADD //Look for units on Balance Sheet table (i.e thousands, millions)'''
    Denomination = df.iloc[2, 0]
    print("Denomination:", Denomination)

    months = df.iloc[3, :]
    print("Months:", months)
    years = df.iloc[4, :]
    print("\nYears:", years)

    LongtermDebt = df[df['one'] == 'Long-term debt']
    print("LongtermDebt:", LongtermDebt)

    LongtermDebt = df.iloc[29, :]
    print("LongtermDebt:", LongtermDebt)

    ShortTermDebt = df.iloc[31, :]
    print("ShortTermDebt:", ShortTermDebt)

    ShareholderEquity = df.iloc[39, :]
    print("ShareholderEquity:", ShareholderEquity)

    ShareholderLoan = None
    Outputlabels = ["Date of Financials used", "Debt", "Long term borrowings", "Short term borrowings", "Equity",
                    "Shareholder's Equity", "Shareholder's loan", \
                    "Final Debt", "Final Equity", "DE ratio"]

    years_average_DE_ratio = None



if __name__ == '__main__':

    #c = pdftables_api.Client(my_pdftables_api_key)
    #c.xlsx(r"C:\Users\Rai Shahnawaz\Desktop\AI Challenge\testCopy.pdf", 'output.xlsx')

    xl_workbook = xlrd.open_workbook('output.xlsx')
    Names=xl_workbook.sheet_names()

    for sheetname in Names:
        if sheetname=="Page 5":
            BalanceSheet=xl_workbook.sheet_by_name(sheetname)

    PageRows=list()
    for rowN in range(BalanceSheet.nrows):
        #print (type(BalanceSheet.row(rowN)))
        #print(BalanceSheet.row(rowN))
        PageRows.append(BalanceSheet.row(rowN))
    labels=['one','two','three','four','five']
    df = pd.DataFrame.from_records(PageRows,columns=labels)

    print (df.head(5))
    len(df.index)

    OutputDF=CalculateDE(df)

    # for sheet in book.sheets():
    #     if (sheet=="Page 7"):
    #         BalanceSheet=sheet
    #     print (sheet)
    #     for row in range(sheet.nrows):
    #         #print (sheet.row(row))
    #         print ("Nothing")

            # Read pdf into DataFrame
    #tabula.convert_into("testCopy.pdf", "output.csv", output_format="csv")

    #df = tabula.read_pdf("testCopy.pdf")
    #tabula.show()