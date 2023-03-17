import re
import openpyxl

# Prompt user for file name
filename = input("Enter the file name: ")

# Open Excel workbook
workbook = openpyxl.load_workbook(filename)

# Get active worksheet
worksheet = workbook.active

# Iterate over rows
for row in worksheet.iter_rows(min_row=2):
    # Check if row has a legal mobile phone number
    phone_numbers = []
    for cell in row:
        if cell.value:
            phone_number = re.findall(r"05[0-9]-?\d{7}", str(cell.value))
            if phone_number:
                phone_numbers.append(phone_number[0])
            else:
                cell.value = ""

    # Remove any extra phone numbers and anything after commas
    if len(phone_numbers) > 1:
        for i, phone_number in enumerate(phone_numbers):
            if "," in phone_number:
                phone_numbers[i] = phone_number.split(",")[0]
        phone_numbers = [phone_numbers[0]]

    # Update row with legal phone number(s)
    if phone_numbers:
        for i, cell in enumerate(row):
            if i == 0:
                cell.value = phone_numbers[0]
            else:
                cell.value = ""

# Save changes
workbook.save(filename)

print("Done!")
