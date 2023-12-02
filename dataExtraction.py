import csv, openpyxl, shutil

csv_file = 'D:/Python_Projects/ExcelManipulations/airtravel.csv'
csv_file2 = 'D:/Python_Projects/ExcelManipulations/airtravel_horizontal.csv'


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
    ws = wb.create_sheet("Horizontal_Data")
    with open(csv_file2) as f:
        reader = csv.reader(f, delimiter=',')   # Careful with the delimiter of your CSV.
        for row in reader:
            ws.append(row)
    wb.save(filename)
    print(f'File created, {filename}')


# Function to read an Excel file and create a second one with the same contents and new name.
def rename_excel(input_file, output_file):
    try:
        # Copy the file without accessing its content
        shutil.copyfile(input_file, output_file)
        print(f"File '{input_file}' successfully copied as '{output_file}'.")

    except Exception as e:
        print(f"An error occurred: {e}")
