# encoding: utf-8

# IMPORT BIB
import sys, os

# CONFIG IMPORT
base_path = os.path.abspath('')
sys.path.append(base_path + "/providers/")
sys.path.append(base_path + "/tmp/")

from providers.Factory import Factory

factory = Factory()

XLSXProvider = factory.openProvider("XLSXProvider")

XLSXProvider.openWorkbook('./tmp/example.xlsx')
print(XLSXProvider.listSheetsName())

sheet = XLSXProvider.openActiveSheet()

# ACESSANDO VALORES
print(sheet['C3'].value)
print(sheet.cell(row=3,column=3).value)

# INSERINDO VALORES
sheet['C3'].value = "Banana"
print(sheet['C3'].value)


# ITERANDO
for row in sheet["B3:D8"]:
    for cell in row:
        print(cell.value)
    print("--- NEW ROW ---")


XLSXProvider.saveWorkbook("./tmp/example_saved.xlsx")


