import pandas as pd
from openpyxl import load_workbook
df_total = pd.DataFrame()
sheets = ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5']
xin = input("Enter gmail")
for sheet in sheets:  # loop through sheets inside an Excel file
    x = pd.read_excel('sample.xlsx', sheet_name=sheet)
    df1 = pd.DataFrame(x[x['gmail']==xin], columns=x.columns)
    df_total = df_total.join(df1, how='outer', lsuffix='_left', rsuffix='_right')
path = "sample.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine='openpyxl')
writer.book = book
if 'MasterSheet' in book.sheetnames:
    ref = book['MasterSheet']
    book.remove(ref)
df_total.to_excel(writer, sheet_name='MasterSheet')
writer.save()
writer.close()
