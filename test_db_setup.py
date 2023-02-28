import sqlite3
import xlrd

db = sqlite3.connect('test_mailback.db')
cur = db.cursor()

try:
    cur.execute("DROP TABLE client;")
except:
    print("client table does not exist")

cur.execute("""CREATE TABLE client(
                full_name       TEXT NOT NULL,
                query_name      TEXT NOT NULL,
                address         TEXT NOT NULL,
                phone_number    TEXT NOT NULL
                )""")

workbook = xlrd.open_workbook('test_clients.xls')
worksheet = workbook.sheet_by_name('mailback')

clients = []
for i in range(1, worksheet.nrows):
    row = worksheet.row_values(i)
    clients += [row]

clients = [tuple(r) for r in clients]

cur.executemany("INSERT INTO client VALUES (?,?,?,?);", clients)



try:
    cur.execute("DROP TABLE mailback_reason;")
except:
    print("mailback_reason table does not exist")


cur.execute("""CREATE TABLE mailback_reason(
                type         TEXT NOT NULL,
                reason       TEXT NOT NULL
                )""")

reasons = [('Date' ,'it is expired',),
           ('Date', 'the date is missing',),
           ('Date', 'the date is illegible'),
           ('Date', 'the date is incorrect',),

           ('Payee', 'the payee is incorrect',),
           ('Payee', 'the payee is illegible'),
           ('Payee', 'the payee is missing',),

           ('Signature', 'the signature is missing',),

           ('Amount', 'the written amounts do not match',),
           ('Amount', 'the written amount is missing'),
           ('Amount', 'the written amount is illegible'),

           ('Other', 'we do not accept this currency',),
           ('Other', 'the check is damaged',),
           ('Other', 'it is marked "Paid in Full"',),
           ('Other', 'there is a restrictive clause on the check',),
           ('Other', 'there is no account number to help us process it',),
           ]

cur.executemany("INSERT INTO mailback_reason VALUES (?, ?);", reasons)

for r in cur.execute("SELECT* FROM client;"):
    print(r)

for r in cur.execute("SELECT* FROM mailback_reason;"):
    print(r)

db.commit()
db.close()
