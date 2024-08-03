from openpyxl import Workbook,load_workbook


wb1 = load_workbook('Vehicle information_20240728_213917.xlsx')
ws1 = wb1.active
print(ws1)