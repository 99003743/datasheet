import os
import os.path
import pandas as pd
from openpyxl import load_workbook
from pandas.tests.io.excel.test_openpyxl import openpyxl

df=[pd.read_excel("Sheet1.xlsx"),
	pd.read_excel("Sheet2.xlsx"),
	pd.read_excel("Sheet3.xlsx"),
	pd.read_excel("Sheet4.xlsx"),
	pd.read_excel("Sheet5.xlsx")]


emp_num = int(input("Enter Employee No : "))


mega_df = []
for i in range(39):
	if(df[1]["number"][i] == emp_num):
		mega_df = [df[0].iloc[[i]],df[1].iloc[[i]],df[2].iloc[[i]],df[3].iloc[[i]],df[4].iloc[[i]]]
result = pd.concat(mega_df, axis=1, join="inner")
result = result.loc[:, ~result.columns.duplicated()]
result = result[result.columns.drop(list(result.filter(regex='Unknown')))]
#print(type(result))
if(os.path.isfile("output.xlsx")):
	path = "output.xlsx"
	book = load_workbook(path)
	writer = pd.ExcelWriter(path, engine='openpyxl')
	writer.book = book
	x = result.iloc[[0]].values.flatten().tolist()


else:
	result.to_excel("output.xlsx", index=False)

