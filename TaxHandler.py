from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, NamedStyle
from datetime import datetime
import TaxHandlerComponents as thc

wb_name = 'ใบวางบิล.xlsx'

wb = load_workbook(wb_name)

while True:
    sheet = thc.ask_for_sheet(wb)

    if not sheet:
        continue

    latest_row_num = thc.find_last_row_with_data(sheet)

    if latest_row_num:
        latest_row = [cell.value for cell in sheet[latest_row_num]]
        print(f"Latest row with data is row number: {latest_row_num}")
        print(f"Data in the latest row: {latest_row}")
    else:
        print("No data found in the sheet.")

    date = []
    date.append(input("Enter Date: \n"))
    date.append(input("Enter Month: \n"))
    date.append(input("Enter Year: \n"))

    amount = int(input("Enter Amount: \n"))

    price = float(input("Enter Price: \n"))
    totalBeforeDiscount = amount * price

    discount = 0.05 * totalBeforeDiscount
    netTotal = totalBeforeDiscount - discount
    vat = netTotal * 0.07
    total = netTotal + vat

    print(f"Total before discount: {totalBeforeDiscount}")
    print(f"Discount: {discount}")
    print(f"Net Total: {netTotal}")
    print(f"VAT: {vat}")
    print(f"Total: {total}")

    new_id = thc.iterate_id(latest_row[0])

    # data in the latest row: ['NW228', date.date(2023, 9, 24, 0, 0), 100, 48, '=C112*D112', 0, '=E112-F112', '=G112*7/100', '=G112+H112', None, None, None, None, None]
    new_row = [new_id, datetime(int(date[2]), int(date[1]), int(date[0])), amount, price, totalBeforeDiscount, discount, netTotal, vat, total]

    # Append the data to row i by writing in cells
    for col_num, value in enumerate(new_row, start=1):
        # Get the cell in the new row
        new_cell = sheet.cell(row=latest_row_num + 1, column=col_num, value=value)

        # Get the corresponding cell in the previous row
        old_cell = sheet.cell(row=latest_row_num, column=col_num)

        # Copy individual styles from the old cell to the new cell
        new_cell.font = Font(name=old_cell.font.name, size=old_cell.font.size, bold=old_cell.font.bold, italic=old_cell.font.italic)
        new_cell.border = Border(left=old_cell.border.left, right=old_cell.border.right, top=old_cell.border.top, bottom=old_cell.border.bottom)
        # new_cell.fill = old_cell.fill
        new_cell.number_format = old_cell.number_format
        # new_cell.protection = old_cell.protection
        # new_cell.number_format = 'DD/MM/YYYY'
        new_cell.alignment = Alignment(horizontal=old_cell.alignment.horizontal, vertical=old_cell.alignment.vertical)


    wb.save(wb_name)
    print(f"Data successfully inserted into {sheet.title} sheet with ID: {new_row[0]}")

    # # Get the sheets by name
    # nw_sheet = wb['NW']
    # tax_sheet = wb['Tax']

    # # Function to calculate totals
    # def calculate_totals(qty, price):
    #     amount = qty * price
    #     discount = 0  # Modify as needed
    #     net_total = amount - discount
    #     vat = net_total * 0.07  # 7% VAT
    #     total = net_total + vat
    #     return amount, discount, net_total, vat, total
    # Name 	Date	Qty	Price	Amt	ส่วนลด	ยอดรวมสุทธิ	แวท	รวม

    # # Input data for NW sheet
    # date = input("Enter Date (dd/mm/yy): ")
    # qty = int(input("Enter Quantity: "))
    # price = float(input("Enter Price: "))

    # # Calculate the amounts
    # amount, discount, net_total, vat, total = calculate_totals(qty, price)

    # # Find the next empty row in NW sheet
    # nw_row = nw_sheet.max_row + 1

    # # Insert data into NW sheet
    # nw_sheet.append([name, date, qty, price, amount, discount, net_total, vat, total])

    # # Input data for Tax sheet
    # tax_id = tax_sheet.max_row + 6700001  # Assuming Tax IDs are continuous and incrementing

    # # Insert data into Tax sheet
    # tax_sheet.append([tax_id, date, name, amount, discount, net_total, vat, total])

    # # Save the workbook with changes
    # wb.save('example_modified.xlsx')

    # print(f"Data successfully inserted into NW and Tax sheets with Tax ID: {tax_id}")
