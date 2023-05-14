import tkinter as tk
from tkinter import ttk
import pyodbc
import pandas as pd
import warnings

# warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

conn_str = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=D:\OneDrive\SongwooDB\songwoo.accdb;"
)
conn = pyodbc.connect(conn_str)
product_codes = '8010-014', '8010-015', '8010-029'

def create_pivot_table(conn=conn, product_codes="test"):
    if not product_codes:
        return

    codes_str = ", ".join(f"'{code}'" for code in product_codes)
    query = f"""
        TRANSFORM Sum(소요량)
        SELECT 품목코드
        FROM tb반제품BOMq
        WHERE 반제품코드 IN ('8010-014', '8010-015', '8010-029')
        GROUP BY 품목코드
        PIVOT 반제품코드;
        """

    result = pd.read_sql(query, conn)

    # Fill missing values with 0
    result = result.fillna(0)
    # print(result)
    return result

def merge_query(product_codes=product_codes):
    bom = create_pivot_table()
    
    query_swerp = "SELECT Item_Code, Item_Name, Model FROM swerp_품목코드_쿼리;"
    swerp = pd.read_sql(query_swerp, conn)
    
    # BOM 테이블과 swerp 테이블을 품목코드(Item_Code)를 기준으로 병합
    merged = bom.merge(swerp, left_on='품목코드', right_on='Item_Code', how='left')
    
    # 필요한 열만 선택
    result = merged[['품목코드', 'Item_Name', 'Model'] + product_codes]
    
    print(result)


# def merge_query(product_codes):
#     bom = create_pivot_table(conn, product_codes)

#     query_swerp = "SELECT Item_Code, Item_Name, Model FROM swerp_품목코드_쿼리;"
#     swerp = pd.read_sql(query_swerp, conn)

#     # BOM 테이블과 swerp 테이블을 품목코드(Item_Code)를 기준으로 병합
#     merged = bom.merge(swerp, left_on='품목코드', right_on='Item_Code', how='left')

#     # 필요한 열만 선택
#     result = merged[['품목코드', 'Item_Name', 'Model'] + product_codes]

#     print(result)

merge_query()
# def on_select():
#     selected_codes = '8010-014', '8010-015', '8010-029'
#     create_pivot_table(conn, selected_codes)

# product_codes = get_selected_product_codes(conn)

# on_select()

# root = tk.Tk()
# root.title("Select Product Codes")

# listbox = ttk.Treeview(root, selectmode="extended")
# listbox["columns"] = ("Product Code",)
# listbox.column("#0", width=0, stretch=tk.NO)
# listbox.column("Product Code", anchor="w", width=100)
# listbox.heading("Product Code", text="Product Code", anchor="w")


# for code in product_codes:
#     listbox.insert("", "end", values=(code,))

# listbox.pack(pady=10)

# select_button = ttk.Button(root, text="Select", command=on_select)
# select_button.pack()

# root.mainloop()

# conn.close()
