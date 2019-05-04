import xlrd
import subprocess

workbook = xlrd.open_workbook("ACUL MASTER Draw Spring 2018.xlsx")
ws = workbook.sheet_by_index(1)
print("Successfully opened workbook...")
file = open("tbodies.html", "r")
print("Successfully opened tbodies.html template...")
tbodies = file.read()
file.close()

dataRows = [15, 21, 17, 23]
for row in dataRows:
    tbodies = tbodies.replace("#E" + str(row), ws.cell(row - 1, 4).value)
    tbodies = tbodies.replace("#I" + str(row), ws.cell(row - 1, 8).value)
    if ws.cell_type(row - 1, 6) == xlrd.XL_CELL_EMPTY: 
        tbodies = tbodies.replace("#G" + str(row) + "and" + "K" + str(row), "")
    else:
        tbodies = tbodies.replace("#G" + str(row) + "and" + "K" + str(row), 
            str(round(ws.cell(row - 1, 6).value))
            + " - " + 
            str(round(ws.cell(row - 1, 10).value)))
    if row == 107: break
    dataRows.append(row + 12)
print("Successfully made replacements in templates...")

file = open("bare.html", "r")
bare = file.read()
print("Successfully opened bare.html template...")
index = bare.replace("<!--subTBodies-->", tbodies)
print("Successfully replaced tbodies into bare...")
file.close()


file = open("index.html", "w")
file.write(index)
print("Successfully wrote new index.html...")
file.close()
print("Completed.")

output = subprocess.call(['commit.sh'])