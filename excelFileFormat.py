from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, GradientFill, Alignment


# Open Workbook and modify it.
def excel_rework(filename):
    wb = load_workbook(filename)
    ws = wb.active
    print('Contents:')
    print(ws)
    # Set the columns' width according to text of each column
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        length = (length + 2) * 1.2
        ws.column_dimensions[column_cells[0].column_letter].width = length

    # Gradient color for one of the cells.
    ws['B5'].fill = fill = GradientFill(stop=("FFFFFF", "000000"))

    wb.save(filename)


# Set format in the first row in which the headers are placed.
def format_header_horizontal(filename):
    wb = load_workbook(filename)
    # Iterate through each worksheet
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Apply the desired format to each cell in the worksheet
        for row in ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.font = Font(name='Arial', size=12, bold=True)
                cell.fill = PatternFill(start_color='7CAAF0', end_color='7CAAF0', fill_type='solid')
    wb.save(filename)


# Set format in the first column in which the headers are placed.
def format_header_vertical(filename):
    wb = load_workbook(filename)
    # Iterate through each worksheet
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Apply the desired format to each cell in the worksheet
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
            for cell in row:
                cell.font = Font(name='Arial', size=12, bold=True)
                cell.fill = PatternFill(start_color='7CAAF0', end_color='7CAAF0', fill_type='solid')
    wb.save(filename)
