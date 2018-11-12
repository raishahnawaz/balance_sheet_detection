import pdftables_api
#import tabula
import xlrd
import pandas as pd
import numpy as np
import glob
import os
import ntpath


path=r"C:\Users\Rai Shahnawaz\Desktop\AI Challenge\reports"
my_pdftables_api_key="cvmm56lxabx7"

def convert_pdfFiles_to_xlsx(src):
    pdfFiles = []
    path = os.path.join(src,"*.pdf")
    for file in glob.glob(path):
        pdfFiles.append(file)
    ExcelFiles = []
    print("PDF Files:", pdfFiles)
    for file in pdfFiles:
        c = pdftables_api.Client(my_pdftables_api_key)
        #c.xlsx(file, file[0:len(file)-4]+'.xlsx')
        ExcelFiles.append(file[0:len(file)-4]+'.xlsx')
    return ExcelFiles

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




def extract_text_excel(files):
    is_table = []
    page_text = []
    page_number = []
    document_number = []
    document_name = []
    path_name = []

    for i, file in enumerate(files):
        try:
            xl = pd.ExcelFile(file)
            for j, sheet in enumerate(xl.sheet_names):
                # save respective document and sheet number
                fileName = ntpath.basename(file)[:-5]
                path_name.append(file)
                document_name.append(fileName)
                document_number.append(i)
                page_number.append(j)
                # convert the sheet to dataframe
                df = xl.parse(sheet)
                # Decision Stump for deciding table or not table
                if df.shape[1] >= 3:
                    # append 1 for table
                    is_table.append(1)
                else:
                    # append 0 for not table
                    is_table.append(0)
                # Convert the dataframe to a single string
                page_string = df.to_string()
                # Remove NaN
                # page_string = page_string.replace('NaN', ' ')
                # Remove Extra Spaces from String
                page_string = ' '.join(page_string.split())
                # Save result in list
                page_text.append(page_string)

        except OSError as e:
            print("Cannot open file : ", file)
        except IOError:
            print('An error occured trying to read the file :', file)
        except xlrd.XLRDError:
            print('Invalid input format for file : ', file)

    page_number = np.array(page_number) + 1
    document_number = np.array(document_number) + 1
    df = pd.DataFrame({'document_number': document_number, 'document_name': document_name,"file_path":path_name, 'sheet_number': page_number,
                       'sheet_text': page_text, 'is_table': is_table})
    return df

def combine_data(generated_data, tagged_data):
    print("generated_data: ", generated_data.shape,"tagged_data", tagged_data.shape)
#    print("generated_data",generated_data.select('document_number', 'document_name'))
    print("generated_data_col: ", generated_data.columns,"tagged_data_col", tagged_data.columns)
    print(generated_data.document_name.unique())

    combined_data = pd.merge(generated_data, tagged_data, left_on=['document_name', 'sheet_number'], right_on=['Document Name', 'Page'])
    print("Length (In Function1): ", combined_data.shape)

    combined_data.drop(['SR#', 'Document', 'Document Name', 'Page'], inplace=True, axis=1)
    combined_data.columns = ['document_number', 'document_name', 'file_path','sheet_number', 'sheet_text', 'is_table_heuristic','is_table_tagged', 'is_balance_sheet']
    combined_data.is_table_tagged = combined_data.is_table_tagged.map({'yes': 1, 'no': 0})
    combined_data.is_balance_sheet = combined_data.is_balance_sheet.map({'yes': 1, 'no': 0})
    return combined_data

if __name__ == '__main__':

    '''READ ALL PDFs FROM A DIRECTORY AND CONVERT ALL FILES INTO XLXS FORMAT USING pdftables-api'''

    check = input('Please enter the folder path where the files are located: ')
    src=path
    if not os.path.isdir(src):
        print('Invalid given path.')
        exit(1)

    XlsxFiles=convert_pdfFiles_to_xlsx(src)
    print("File Names : ", XlsxFiles)


    generated_data=extract_text_excel(XlsxFiles)
    print("Generated Data : ", generated_data.shape)

    print("TablePages[0]", generated_data.document_name.unique())


    try:
        tagged_data = pd.read_csv('Tagged Data For AI Chellange - Sheet1.csv')
    except OSError as e:
        print("Cannot open file : ")
    except IOError:
        print('An error occured trying to read the file :')
    except xlrd.XLRDError:
        print('Invalid input format for file : ')

    print("Tagged Data : ", tagged_data.shape)

    combined_data=combine_data(generated_data, tagged_data)

    print("Length (combined_data): ", combined_data.shape)
    print (combined_data.head(5))

    '''OPEN A CONVERTED REPORT AND LOOK FOR THE BALANCE SHEET'''
    filename='output.xlsx'
    Balance_sheet_DF=load_balance_sheet_toDF(filename)

    '''BALANCE SHEET PARSING FOR DEBT EQUITY VALUES'''
    DE_output_DF=parse_balance_sheet_for_DE(Balance_sheet_DF)

    '''Calculating all the required fields using debt equity values'''
    OutputDF=CalculateDE(DE_output_DF)

