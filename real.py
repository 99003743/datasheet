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


emp_num = list(map(str, input("Enter Employee No's : ").split(" ")))

mega_df = []
result1 = []
result = 0
for j in emp_num:
	for i in range(39):
		if(str(df[1]["PsNo"][i]) == j):
			mega_df = [df[0].iloc[[i]],df[1].iloc[[i]],df[2].iloc[[i]],df[3].iloc[[i]],df[4].iloc[[i]]]

			#Merging columns data from multiple sheets
			result = pd.concat(mega_df, axis=1, join="inner")

			#Removing duplicate columns which are name, emp id, gmail
			result = result.loc[:, ~result.columns.duplicated()]

			#Removing any additional duplicate columns which are formed accidentally by using regular expressions
			result = result[result.columns.drop(list(result.filter(regex='Unknown')))]

			#Appending this row data into a list
			result1.append(result.iloc[[0]].values.flatten().tolist())

#Creating DataFrame from the row data list
result2 = pd.DataFrame(result1, columns=result.columns)

#Storing dataframe into an excel sheet
result2.to_excel("output.xlsx", index=False)
