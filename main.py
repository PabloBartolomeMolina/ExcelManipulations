import openpyxl
import csv
from openpyxl import Workbook
import time

name = 'D:/airtravel.xlsx'


def read_wb():
    wb = openpyxl.Workbook()
    ws = wb.active
    # CSV file can be downloaded from "https://people.sc.fsu.edu/~jburkardt/data/csv/csv.html"
    with open('D:/airtravel.csv') as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
            ws.append(row)
    wb.save(name)
    print(f'File created, {name}')


if __name__ == '__main__':
    read_wb()
