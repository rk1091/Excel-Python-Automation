'''
python, openpyxl
'''

#assuming both input output files exist (later remove that too)
#make an array and input name and other values 0,1 in that and append 
# that in another sheet ie output

import openpyxl

excel_files = ['G:\rish project\try1\INPUT.xlsx']
output_file=['G:\rish project\try1\OUTPUT.xlsx']
#if (worksheet['E19']>=1649 and worksheet['E19']<=1651)


for file in excel_files:
    wb=openpyxl.load_workbook(file)
    worksheet= wb["Sheet1"]
    a=[]
    a[0]=worksheet['C13']
    w=openpyxl.load_workbook(output_file)
    worksheet= w["Sheet1"]
    w.append(a)


    
    