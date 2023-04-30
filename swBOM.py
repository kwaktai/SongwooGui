import pyodbc
import pandas as pd
from tkinter import *
from tkinter import ttk

# 데이터베이스 연결
def connect_to_db(file_path):
    connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file_path
    return pyodbc.connect(connection_string)

# 반제품코드 목록 가져오기
def get_product_codes(conn):
    query = "SELECT DISTINCT 반제품코드 FROM tb반제품BOMq;"
    return pd.read_sql(query, conn)['반제품코드'].tolist()

# 피벗 테이블 생성 및 저장
def create_pivot_table(conn, product_codes):
    if not product_codes:
        return

    codes_str = ", ".join(f"'{code}'" for code in product_codes)
    query = f"""
        TRANSFORM Sum(소요량)
        SELECT 품목코드
        FROM tb반제품BOMq
        WHERE 반제품코드 IN ({codes_str})
        GROUP BY 품목코드
        PIVOT 반제품코드;
        """

    result = pd.read_sql(query, conn)

    # Fill missing values with 0
    result = result.fillna(0)

    # Save the result as an Excel file
    result.to_excel("output.xlsx", index=False)
    # 'print("codes_str: " + codes_str)
    # print("result: " + result)


# 반제품코드 선택 팝업
def select_product_codes(conn):
    def on_select():
        selected_items = listbox.curselection()
        selected_codes = [listbox.get(index) for index in selected_items]
        create_pivot_table(conn, selected_codes)
        root.destroy()

    root = Tk()
    root.title("반제품코드 선택")

    listbox = Listbox(root, selectmode="extended")

    product_codes = get_product_codes(conn)
    for code in product_codes:
        listbox.insert(END, code)

    listbox.pack(fill=BOTH, expand=1)

    btn_select = Button(root, text="선택", command=on_select)
    btn_select.pack(side=RIGHT)

    btn_cancel = Button(root, text="취소", command=root.destroy)
    btn_cancel.pack(side=LEFT)

    root.mainloop()

if __name__ == "__main__":
    file_path = "D:\OneDrive\SongwooDB\songwoo.accdb"
    conn = connect_to_db(file_path)
    select_product_codes(conn)





         품목코드   Item_Name                         Model  8010-030  8010-037  8010-051  8010-072
0    AMP-0001    IC-OPAMP                   KA358A/SOP8       0.0       1.0       0.0       0.0
1    BAT-0016  BAT-BACKUP                CR2032/3PIN/ST       0.0       1.0       0.0       0.0


          품목코드   Item_Name                         Model  PS-130S_MAIN_SN-100  MS-233_MAIN(SK)  EP-642_MAIN BOARD  ED-643N_MAIN
0    AMP-0001    IC-OPAMP                   KA358A/SOP8       0.0       1.0       0.0       0.0
1    BAT-0016  BAT-BACKUP                CR2032/3PIN/ST       0.0       1.0       0.0       0.0
