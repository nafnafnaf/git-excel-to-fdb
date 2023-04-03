import xlrd
import fdb
wb = xlrd.open_workbook('TEST-2.xlsx')
sheet = wb.sheet_by_index(0)
# print("Value of 20-4 cell: ",sheet.cell_value(2, 3))
# print("Number of Rows: ", sheet.nrows)
# print("Number of Columns: ",sheet.ncols)
# print("ALL COLUMN NAMES ARE: ")
# for i in range(sheet.ncols):
#     print(sheet.cell_value(0,i))
rows = sheet.nrows
cols = sheet.ncols
print(f"ΠΛΗΘΟΣ ΓΡΑΜΜΩΝ: {rows} ΠΛΗΘΟΣ ΣΤΗΛΩΝ {cols} ")
#print(f"sunolo seirwn:{values}")
#colidx = dict( (sheet.cell(0, i).value, i) for i in range(sheet.ncols) )
# oles oi times apo tin stili 0
# val = sheet.col_values(0)
# print(val)
# k = []
# for el in val:
#     k.append(str(el))
# print(type(k))
# print(len(k))


# for row in range(sheet.nrows):
#     values = []
#     for col in range(sheet.ncols):
#        values.append(sheet.cell(row,col).value)
#     z = values
#     print(z)

# ls = [str(wb.sheet_by_index(0).cell_value(i,0)) for i in range(wb.sheet_by_index(0).nrows)]
# ls=[list(map(int,i.split(' '))) for i in ls]
# print(ls)


# n = 9
# x = [values[i:i + n] for i in range(0, len(values), n)]
# print(x)
#lista = [(17, 'rachel', 67), (18, 'ross', 79), (19, 'nick', 95), (19, 'nick', 95)]


# # graph = [tuple(map(int, i.split(','))) for i in k] #REMOVE COMMA
# # print(graph)
# w = [f'({x})' for x in k] # ADD PARENTH.
# n = 9
# x = [w[i:i + n] for i in range(0, len(w), n)]
#
# # t = [i.strip('"') for i in x]
# z = x[3]
# g1 = [i.replace('', '') for i in z]
# print(len(z))
# print(g1)
# #l = [x.strip("'") for x in g1]
#
# print(g1)
# #records_list = [('48'), (19997), (199), (19994), (19), (193399), (9994), (999)]# SWSTH MORFH LISTAS MAZI ME cur.execute(insert_query, records_list)
database='D:\FIREBIRD DBS\FARMA-CY.FDB'
con = fdb.connect(host='localhost',
                      database=database,
                      user='sysdba',
                      password='dimi',
                  )

if con:
    print(f'ΣΥΝΔΕΘΗΚΕ ΜΕ ΒΑΣΗ! {database}')

cur = con.cursor()
insert_query = """INSERT INTO SKEYASMATA
               (DRUGCODE, NAME, VATINDEX, COSTPRICE, RETAILPRICE, APOTHNAME,  MNSTRCODE, BARCODE)
               VALUES (?,?,?,?,?,?,?,?)"""

for row in range(sheet.nrows):
    values = []
    for col in range(sheet.ncols):
       values.append(sheet.cell(row,col).value)
    cur.execute(insert_query, values)
con.commit()
