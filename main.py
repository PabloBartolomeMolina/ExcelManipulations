import excelFileFormat, dataExtraction


# Main entry point.
if __name__ == '__main__':
    dataExtraction.csv_to_excel()
    excelFileFormat.excel_rework()
    excelFileFormat.format_header()
