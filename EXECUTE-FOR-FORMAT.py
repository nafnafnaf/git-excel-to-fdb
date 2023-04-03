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
#colidx = dict( (sheet.cell(0, i).value, i) for i in range(sheet.ncols) )
# oles oi times apo tin stili 0
val = sheet.col_values(0)
print(len(val))
print(type(val))
print(val[4])
k = []
for el in val:
    k.append(str(el))
# print(type(k))
# print(k)
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

#u = [f'({i})' for i in z]
#for i in u:
#    i.replace('"', '')
# print(u)

# length = len(k)
# for a in range(length):

#for a in range(4):
     # print(z[a])
#for i in z:
# a = [1,2,3,4,5,6]
for i in range(len(z)):
    # con.execute(insert_query.replace('%DRUGCODE%', z[i], 1 ))
    con.execute_immediate(insert_query.format(z[i], i ))
#
# # cur.execute(insert_query, z)
con.commit()
# # Retrieve all rows as a sequence and print that sequence:
# print(cur.fetchall())
# print(type(cur))



"Δεν είναι δυνατό να γίνει εξερχόμενη κλήση επειδή η εφαρμογή στέλνει μία κλήση σύγχρονης εισόδου"
An outgoing call cannot be made since the application is dispatching an input-synchronous call