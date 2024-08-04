import os
import pandas as pd
from openpyxl import Workbook,load_workbook

#New sheet with openpyxl
wb1 = load_workbook('Vehicle information_20240728_213917.xlsx')#copy paste the new generated sheet from DMS to here
#ws1 = wb1.active
check_sheet = wb1.sheetnames
if check_sheet.count('sorted')==0: 
    newWS = wb1.create_sheet(title="sorted")#create a new sheet 'sorted'


#select the columns with pandas
xlsheet = pd.read_excel('Vehicle information_20240728_213917.xlsx', sheet_name='Vehicle information')#copy paste the new generated sheet from DMS to here
copiedColumn = [3, 9, 15, 23] #column of 'VIN','Model','Delivery date', 'Target country for vehicle sales'
selected_column = xlsheet.iloc[:, copiedColumn]

wb1.save('Vehicle information_20240728_213917.xlsx')
#print(check_sheet)