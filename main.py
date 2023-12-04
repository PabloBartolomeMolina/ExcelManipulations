import dataExtraction
import excelFileFormat

excel_file = 'D:/Python_Projects/ExcelManipulations/airtravel.xlsx'


# Main entry point.
if __name__ == '__main__':
    dataExtraction.csv_to_excel(excel_file)
    excelFileFormat.excel_rework(excel_file)
    excelFileFormat.format_header_vertical(excel_file)
