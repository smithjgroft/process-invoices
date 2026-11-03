import pandas as pd
from analyticskit import *
from datetime import datetime as dt
import openpyxl


while True:
    invoice_data_path = select_file('select invoice file')
    if "Invoice-Data.xlsx" in invoice_data_path:
        break
    else:
        show_message('error', 'incorrect file')


while True:
    invoice_template_path = select_file('select invoice template file')
    if "Invoice-Template.xlsx" in invoice_template_path:
        break
    else:
        show_message('error', 'incorrect file')

directory = invoice_template_path.split('Invoice-Template.xlsx')[0]



invoice_data = pd.read_excel(invoice_data_path)

for index,row in invoice_data.iterrows():
    
    invoice_template = openpyxl.load_workbook(invoice_template_path)
    
    selected_sheet = invoice_template['ACCInvoice']
    
    selected_sheet['C3'].value = 'Advanced Accounting Analytics'
    selected_sheet['C4'].value = '1234 Python St., Provo, UT 84602'
    selected_sheet['C5'].value = '(801) 555-0123'
    selected_sheet['C6'].value = 'pythonmaster@byu.edu'
    
    selected_sheet['G6'].value = dt.today().strftime('%m/%d/%Y')
    selected_sheet['G8'].value = row['Invoice Number']
    
    selected_sheet['B14'].value = row['Name']
    selected_sheet['B15'].value = row['Company Name']
    selected_sheet['B16'].value = f"{row['Street Address']}, {row['City']}, {row['State']} {row['Zip Code']}"
    selected_sheet['B17'].value = row['Phone']
    selected_sheet['B18'].value = row['Email Address']
    
    if pd.isna(row['Item1']) == False:
        selected_sheet['B22'].value = row['Item1']
        selected_sheet['E22'].value = row['Quantity1']
        selected_sheet['F22'].value = row['Price1']
        
    if pd.isna(row['Item2']) == False:
        selected_sheet['B23'].value = row['Item2']
        selected_sheet['E23'].value = row['Quantity2']
        selected_sheet['F23'].value = row['Price2']
        
    if pd.isna(row['Item3']) == False:
        selected_sheet['B24'].value = row['Item3']
        selected_sheet['E24'].value = row['Quantity3']
        selected_sheet['F24'].value = row['Price3']
        
    invoice_template.save(f"{directory}Invoice_{row['Invoice Number']}.xlsx")