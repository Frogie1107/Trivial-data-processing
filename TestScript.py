import os
import pandas as pd
from openpyxl import Workbook,load_workbook

#Check if new sheet was made.If not, new sheet with openpyxl
wb1 = load_workbook('Vehicle information_20240728_213917.xlsx')#copy paste the DMS new generated sheet's name to here
#ws1 = wb[''] #access different worksheet
check_sheet = wb1.sheetnames
if check_sheet.count('sorted')==0: 
    ws = wb1.create_sheet(title="sorted")#create a new sheet call 'sorted'
#print(wb1.sheetnames) #check if the worksheet are there


#select the columns with pandas
xlsheet = pd.read_excel('Vehicle information_20240728_213917.xlsx', sheet_name='Vehicle information')#read the excel sheet 'Vehicle information' from excel file
copiedColumn = [3, 9, 15, 23] #column of 'VIN','Vehicle series code','Delivery date', 'Target country for vehicle sales' from excel sheet
selected_column = xlsheet.iloc[:, copiedColumn]


#wb1.save('Vehicle information_20240728_213917.xlsx')# make changes on the excel file