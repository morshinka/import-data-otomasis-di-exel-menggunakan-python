import pymysql 
import xlwings as xw
import time

connetion = pymysql.connect(
    host="127.0.0.1",
    user='root',
    password='',
    database='db_name'
    
)

file_path = "fill local directory microsoft excel file"

app = xw.App(visible=True)
wb = app.books.open(file_path)
ws = wb.sheets[0]

try:
    with connetion.cursor() as cursor:
        query = "SELECT * FROM users LIMIT 20"
        cursor.execute(query)
        
        result = cursor.fetchall()
        
        start_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
        
        for i, row in enumerate(result, start=start_row):
            ws.range(f'A{i}').value = row[0]
            ws.range(f'B{i}').value = row[1]
            ws.range(f'C{i}').value = row[2]
            print(f'input data baris ke-{i}: number: {row[0]}, name: {row[1]}, email: {row[2]}')
            time.sleep(1)
finally:
    wb.save()
    connetion.close()
