import vcfToEcxel

excel_file = 'D:/Python_Projects/ExcelManipulations/airtravel.xlsx'


# Main entry point.
if __name__ == '__main__':
    '''
    dataExtraction.csv_to_excel(excel_file)
    excelFileFormat.excel_rework(excel_file)
    excelFileFormat.format_header_vertical(excel_file)
    '''
    vcf_file_path = 'D:/FileWithContacts.vcf'
    excel_file_path = 'D:/Ordered_Contacts.xlsx'
    vcfToEcxel.vcf_to_excel(vcf_file_path, excel_file_path)
