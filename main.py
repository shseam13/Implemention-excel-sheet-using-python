"""All rights reserved by @shseam13
follow me on-
    Github:    https://github.com/shseam13
    Facebook:  https://www.facebook.com/shajjadhossains1/
    LinkedIn:  https://bd.linkedin.com/in/shajjad-hossain-seam-b6ba641b0
"""
from random import randint
import openpyxl

workbook = openpyxl.Workbook()
sheet =workbook.active

data = (
    ("Customer", "RN for IAT", "IAT", "TA", "RN for ST", "ST", "TSB", "WT", "TSE", "T Spent in Sys", "Idle TOS"),
)

for row in data:
    sheet.append(row)

# below column --> Customer number
for val in range(2,8):
    column_number = f"A{val}"
    customer = sheet[column_number]
    customer.value = val-1

# below column --> Random number for IAT
#--> static code:
sheet["B2"].value = "--"
sheet["B3"].value = 207
sheet["B4"].value = 127
sheet["B5"].value = 19
sheet["B6"].value = 559
sheet["B7"].value = 908
#--> dynamic code:
# for val in range(3,8):
#     column_number = f"B{val}"
#     sheet["C2"].value = 0
#     rn_iat = sheet[column_number]
#     rn_iat.value = randint(1,1000)

# below column - IAT
for val in range(3,8):
    column_number = f"C{val}"
    column_num_previous = f"B{val}"
    b = sheet[column_num_previous]
    c = sheet[column_number]
    if b.value < 125:
        c.value = 1
    elif b.value < 250:
        c.value = 2
    elif b.value < 375:
        c.value = 3
    elif b.value < 500:
        c.value = 4
    elif b.value < 625:
        c.value = 5
    elif b.value < 750:
        c.value = 6
    elif b.value < 875:
        c.value = 7
    elif b.value < 1000:
        c.value = 8

# below column - Arrival Time
for val in range(2,7):
    sheet["D2"].value  = 0 
    sheet[f"D{val+1}"].value = (sheet[f"D{val}"].value) + (sheet[f"C{val+1}"].value)

# below column - Random Service Time
#--> static code:
sheet["E2"].value = 95
sheet["E3"].value = 2
sheet["E4"].value = 55
sheet["E5"].value = 82
sheet["E6"].value = 15
sheet["E7"].value = 42
#--> dynamic code
# for val in range(2,8):
#     sheet[f"E{val}"].value = randint(1,100)

# below column - Service Time
for val in range(2,8):
    column_number = f"F{val}"
    column_num_previous = f"E{val}"
    b = sheet[column_num_previous]
    c = sheet[column_number]
    if b.value < 21:
        c.value = 1
    elif b.value < 31:
        c.value = 2
    elif b.value < 61:
        c.value = 3
    elif b.value < 71:
        c.value = 4
    elif b.value < 96:
        c.value = 5
    elif b.value < 101:
        c.value = 6

# below column - Time service begins and Time Service ends
sheet["I2"].value = sheet["F2"].value
sheet["G2"].value = 0
for val in range(3,8):
    sheet[f"G{val}"].value = max(sheet[f"D{val}"].value, sheet[f"I{val-1}"].value)
    sheet[f"I{val}"].value = sheet[f"F{val}"].value + sheet[f"G{val}"].value

# below column - Waiting Time
for val in range(2,8):
    sheet[f"H{val}"].value = abs(sheet[f"G{val}"].value - sheet[f"D{val}"].value)

# below column - Time Spent in System
for val in range(2,8):
    sheet[f"J{val}"].value = abs(sheet[f"I{val}"].value - sheet[f"D{val}"].value)

# below column - Time Spent in System
for val in range(3,8):
    sheet[f"K{val}"].value = abs(sheet[f"I{val-1}"].value - sheet[f"G{val}"].value)

# save your workbook
workbook.save(filename="output.xlsx")
