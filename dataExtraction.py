import csv, openpyxl

csv_file = 'D:/Python_Projects/ExcelManipulations/airtravel.csv'


# Function to read a CSV and create an Excel ".xlsx" file with the contents.
def csv_to_excel(filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Raw data'
    # CSV file can be downloaded from "https://people.sc.fsu.edu/~jburkardt/data/csv/csv.html"
    with open(csv_file) as f:
        reader = csv.reader(f, delimiter=',')   # Careful with the delimiter of your CSV.
        for row in reader:
            ws.append(row)
    wb.save(filename)
    print(f'File created, {filename}')