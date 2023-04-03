import xlrd
import fdb
wb = xlrd.open_workbook('TEST-2.xlsx')
sheet = wb.sheet_by_index(0)
rows = sheet.nrows
cols = sheet.ncols
print(f"ΠΛΗΘΟΣ ΓΡΑΜΜΩΝ: {rows} ΠΛΗΘΟΣ ΣΤΗΛΩΝ {cols} ")

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
