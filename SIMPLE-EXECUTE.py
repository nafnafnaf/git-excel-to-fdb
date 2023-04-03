import fdb

records_list = [('4805788'), (4404444), (1909), (109994), (19), (193399), (9994), (999)]


con = fdb.connect(host='localhost',
                      database='D:\FIREBIRD DBS\FARMA-CY.FDB',
                      user='sysdba',
                      password='dimi',
                  )
if con:
    print('YES!')
    print(len(records_list),type(records_list))
cur = con.cursor()
insert_query = """INSERT INTO SKEYASMATA
               (BARCODE, DRUGCODE, VATINDEX, NAME, MNSTRCODE, APOTHNAME, COSTPRICE, RETAILPRICE)
               VALUES (?,?,?,?,?,?,?,?)"""
cur.execute(insert_query, records_list)
con.commit()