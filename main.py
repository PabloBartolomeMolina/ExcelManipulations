import openpyxl
import csv

from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import time

excel_file = 'D:/Python_Projects/ExcelManipulations/airtravel.xlsx'


# Function to read a CSV and create an Excel ".xlsx" file with the contents.
def csv_to_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Raw data'
    # CSV file can be downloaded from "https://people.sc.fsu.edu/~jburkardt/data/csv/csv.html"
    with open('D:/Python_Projects/ExcelManipulations/airtravel.csv') as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
            ws.append(row)
    wb.save(excel_file)
    print(f'File created, {excel_file}')


# Open Workbook and modify it.
def excel_rework():
    wb = load_workbook(excel_file)
    ws = wb.active
    print('Contents:')
    print(ws)

    ws['A1'].fill = PatternFill("solid", fgColor="aa5624")
    #ws['A1'].fill = fill = GradientFill(stop=("000000", "FFFFFF"))
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

    #ws['D1'].fill = PatternFill("solid", fgColor="aa5624")
    ws['D1'].fill = fill = GradientFill(stop=("FFFFFF", "000000"))
    ws['D1'].alignment = Alignment(horizontal="center", vertical="center")

    wb.save(excel_file)

# Main entry point.
if __name__ == '__main__':
    csv_to_excel()
    excel_rework()
