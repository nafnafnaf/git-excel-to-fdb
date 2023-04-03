import xlrd
import fdb
wb = xlrd.open_workbook('MEDAPTH-HOUVAR-060922.xlsx')
sheet = wb.sheet_by_index(0)
print("Value of 20-4 cell: ",sheet.cell_value(20, 4))
print("Number of Rows: ", sheet.nrows)
print("Number of Columns: ",sheet.ncols)
print("ALL COLUMN NAMES ARE: ")
for i in range(sheet.ncols):
    print(sheet.cell_value(0,i))

rows= sheet.nrows
cols = sheet.ncols
print(f"There are {rows} rows and {cols} columns in this sheet")
for row in range(sheet.nrows):
    values = row
print(f"sunolo seirwn:{values}")

val = sheet.col_values(0)
print(len(val))
print(type(val))
print(val[4])
k = []
for el in val:
    k.append(str(el))

con = fdb.connect(host='localhost',
                      database='D:\FIREBIRD DBS\FARMA-CY.FDB',
                      user='sysdba',
                      password='dimi',
                  )
if con:
    print('YES!')

cur = con.cursor()

insert_query = """INSERT INTO SKEYASMATA
               (DRUGCODE, VATINDEX)
               VALUES ('{}', '{}' )"""
              # VALUES('%DRUGCODE%', '%vatindex$')
z = k

for i in range(len(z)):

    con.execute_immediate(insert_query.format(z[i], i ))  #PROSOXH STH FORMAT

con.commit()



