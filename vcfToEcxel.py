import vobject
import openpyxl


def vcf_to_excel(vcf_file, excel_file):
    # Load VCF file
    with open(vcf_file, 'r', encoding='utf-8') as vcf_data:
        vcard_list = []
        current_vcard = ''
        for line in vcf_data:
            # Skip lines containing only ";;;"
            if line.strip() == ';;;':
                continue

            current_vcard += line

            # Check if the line ends a vCard
            if line.strip() == 'END:VCARD':
                vcard_list.append(current_vcard)
                current_vcard = ''
    # Sort the vcard_list based on the Name (first column)
    vcard_list.sort(key=lambda vcard_data: vobject.readOne(vcard_data).fn.value.lower() if hasattr(vobject.readOne(vcard_data), 'fn') else '')
    # Create Excel workbook and sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write headers
    headers = ['Name', 'Email', 'Phone']
    sheet.append(headers)

    # Extract data from VCF and write to Excel
    for vcard_data in vcard_list:
        vcard = vobject.readOne(vcard_data)

        name = vcard.fn.value if hasattr(vcard, 'fn') and vcard.fn.value else ''
        email = vcard.email.value if hasattr(vcard, 'email') and vcard.email.value else ''
        phone = vcard.tel.value if hasattr(vcard, 'tel') and vcard.tel.value else ''

        sheet.append([name, email, phone])

    # Format the header row (bold and centered)
    for cell in sheet[1]:
        cell.font = openpyxl.styles.Font(bold=True)
        cell.alignment = openpyxl.styles.Alignment(horizontal='center')
    # Adjust column widths
    for column in sheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = adjusted_width

    # Save Excel file before closing the workbook
    workbook.save(excel_file)
