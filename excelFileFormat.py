from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, GradientFill, Alignment

excel_file = 'D:/Python_Projects/ExcelManipulations/airtravel.xlsx'


# Open Workbook and modify it.
def excel_rework():
    wb = load_workbook(excel_file)
    ws = wb.active
    print('Contents:')
    print(ws)

    # Format for 1st row : Font size = 12 and Bold letters
    # Set a background color to this first row
    # Set centered alignment to this first row
    for column_cell in ws.iter_cols(1, 1):  # iterate column cells
        for row_cell in ws.iter_rows(1, ws.max_row + 1):
            for cell in column_cell:
                cell.font = Font(b=True, size=12)
                cell.fill = PatternFill(start_color="7CAAF0", end_color="7CAAF0", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Set the columns' width according to text of each column
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        length = (length + 2) * 1.2
        ws.column_dimensions[column_cells[0].column_letter].width = length

    # Gradient color for one of the cells.
    ws['A5'].fill = fill = GradientFill(stop=("FFFFFF", "000000"))

    wb.save(excel_file)


def format_header():
    wb = load_workbook(excel_file)
    # Iterate through each worksheet
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Apply the desired format to each cell in the worksheet
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.font = Font(name='Arial', size=12, bold=True)
                cell.fill = PatternFill(start_color='7CAAF0', end_color='7CAAF0', fill_type='solid')
