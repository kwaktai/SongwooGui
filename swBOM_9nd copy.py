import pyodbc

def get_product_codes_and_models():
    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\OneDrive\SongwooDB\songwoo.accdb;')
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT 완제품코드, Model FROM tb완제품BOMq_반제품기준')
    rows = cursor.fetchall()
    conn.close()
    return rows

product_codes_and_models = get_product_codes_and_models()

for row in product_codes_and_models:
    print(f"Product Code: {row[0]}, Product Model: {row[1]}")
