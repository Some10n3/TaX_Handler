def find_last_row_with_data(sheet):
    # Loop through rows from the bottom up
    for row in range(sheet.max_row, 0, -1):
        # Check if any cell in the row has a value
        if any(cell.value for cell in sheet[row]):
            return row
    return None

def ask_for_sheet(wb):
    sheet = input("Enter sheet: ")

    if sheet not in wb.sheetnames:
        print(f"Sheet '{sheet}' not found in the workbook.")
        return None
    else:
        print("Sheet selected: ", sheet)
        return wb[sheet]

def iterate_id(id):
    # 'NW228' -> 'NW229'
    # SPS001 -> SPS002

    # Extract the prefix (letters) and the numeric part
    prefix = ''.join(filter(str.isalpha, id))  # All alphabetic characters
    number = ''.join(filter(str.isdigit, id))  # All numeric characters

    # Convert the numeric part to an integer and increment it
    new_number = int(number) + 1

    # Format the new ID, preserving the original number of digits in the numeric part
    return f"{prefix}{new_number:0{len(number)}d}"

