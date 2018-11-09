import pdftables_api
#import tabula
import xlrd
import pandas as pd
import numpy


if __name__ == '__main__':

    #c = pdftables_api.Client(my_pdftables_api_key)
    #c.xlsx(r"C:\Users\Rai Shahnawaz\Desktop\AI Challenge\testCopy.pdf", 'output.xlsx')

    #path="C:\Users\Rai Shahnawaz\Desktop\AI Challenge\\reports\"

    import glob

    txtfiles = []
    for file in glob.glob("*.pdf"):
        txtfiles.append(file)

    print ("txtfiles:", txtfiles)

