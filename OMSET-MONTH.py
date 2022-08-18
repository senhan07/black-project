
# import module
from lib2to3.pgen2.pgen import DFAState
from os import system, name
from traceback import print_stack
from openpyxl import load_workbook
import pandas as pd
import numpy as np
import os
pd.options.mode.chained_assignment = None

pwd = os.getcwd()

def clear():
    if name == 'nt':
        _ = system('cls')
    else:
        _ = system('clear')
clear()



print("OMSET SALES EXTRACT TOOLS")
#filename = input("File Excel:")
filename = "C:/Users/Lenovo/Desktop/data.xls"

sales_array = [
    'P.MAR'
]

jurusan_data = "UTARA"

for sales_data in sales_array:
		sales = pd.read_excel(filename, skiprows=[0], usecols = "B,H,I,J,Y")
		jurusan = pd.read_excel(filename, skiprows=[0], usecols = "B,H,I,J,Y")
		
		
		
		sales = sales[sales['Salesman'].str.contains(sales_data, na=False)]
		jurusan = sales[sales['Area'].str.contains(jurusan_data, na=False)]
		
		#BULAN 1
		tanggal = jurusan[jurusan['Tanggal'].str.contains('1/2022', na=False)]
		tanggal['Total'] = tanggal.groupby(['Customer'])['Total'].transform('sum')
		df1 = tanggal.drop_duplicates(subset=['Customer'])
		df1 = df1.sort_values('Customer')
		total = df1['Total'].sum()
		print(df1)
		
		#BULAN 2
		tanggal = jurusan[jurusan['Tanggal'].str.contains('2/2022', na=False)]
		tanggal['Total'] = tanggal.groupby(['Customer'])['Total'].transform('sum')
		df2 = tanggal.drop_duplicates(subset=['Customer'])
		df2 = df2.sort_values('Customer')
		total = df2['Total'].sum()
		print(df2)

		#BULAN 3
		tanggal = jurusan[jurusan['Tanggal'].str.contains('3/2022', na=False)]
		tanggal['Total'] = tanggal.groupby(['Customer'])['Total'].transform('sum')
		df3 = tanggal.drop_duplicates(subset=['Customer'])
		df3 = df3.sort_values('Customer')
		total = df3['Total'].sum()
		print(df3)
		
		#BULAN 4
		tanggal = jurusan[jurusan['Tanggal'].str.contains('4/2022', na=False)]
		tanggal['Total'] = tanggal.groupby(['Customer'])['Total'].transform('sum')
		df4 = tanggal.drop_duplicates(subset=['Customer'])
		df4 = df4.sort_values('Customer')
		total = df4['Total'].sum()
		print(df4)
	
		#BULAN 5
		tanggal = jurusan[jurusan['Tanggal'].str.contains('5/2022', na=False)]
		tanggal['Total'] = tanggal.groupby(['Customer'])['Total'].transform('sum')
		df5 = tanggal.drop_duplicates(subset=['Customer'])
		df5 = df5.sort_values('Customer')
		total = df5['Total'].sum()
		print(df5)
	
		#BULAN 6
		tanggal = jurusan[jurusan['Tanggal'].str.contains('6/2022', na=False)]
		tanggal['Total'] = tanggal.groupby(['Customer'])['Total'].transform('sum')
		df6 = tanggal.drop_duplicates(subset=['Customer'])
		df6 = df6.sort_values('Customer')
		total = df6['Total'].sum()
		print(df6)
		
		frames = [df1, df2, df3, df4, df5, df6, ]
		result = pd.concat(frames)
		print(result)
		
		# Create a Pandas Excel writer using XlsxWriter as the engine.
		writer = pd.ExcelWriter('C:/Users/Lenovo/Desktop/MarksData.xlsx', engine='xlsxwriter')
		result.to_excel(writer, sheet_name='Sheet1')
		
		
		result['Total'] = result.groupby(['Customer'])['Total'].transform('sum')
		nama_toko = result.drop_duplicates(subset=['Customer'])
		nama_toko = nama_toko.sort_values('Customer')
		
		# Write each dataframe to a different worksheet.
		nama_toko.to_excel(writer, sheet_name='Sheet2')
		
		# Close the Pandas Excel writer and output the Excel file.
		writer.save()
		
		#file_name = 'C:/Users/Lenovo/Desktop/MarksData1.xlsx'
		#result.to_excel(file_name, sheet_name='Sheet1')
		#nama_toko.to_excel(file_name, sheet_name='Sheet2')
		print('DataFrame is written to Excel File successfully.')
	