import tkinter as tk
from tkinter import ttk
import pyodbc
import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

conn_str = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=D:\OneDrive\SongwooDB\songwoo.accdb;"
)
conn = pyodbc.connect(conn_str)


def get_selected_product_codes(conn):
    query = """
    SELECT tb반제품BOMq.반제품코드, swerp_품목코드_쿼리.Model
    FROM tb반제품BOMq
    INNER JOIN swerp_품목코드_쿼리 ON tb반제품BOMq.반제품코드 = swerp_품목코드_쿼리.Item_Code
    GROUP BY tb반제품BOMq.반제품코드, swerp_품목코드_쿼리.Model;
    """
    return pd.read_sql(query, conn)


def create_pivot_table(conn, product_codes):
    if not product_codes:
        return

    codes_str = ", ".join(f"'{code}'" for code in product_codes)
    # print(codes_str)
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

    return result

def merge_query(product_codes):
    orderBOM = create_pivot_table(conn, product_codes)

    query_swerp = "SELECT Item_Code, Item_Name, Model FROM swerp_품목코드_쿼리;"
    swerp = pd.read_sql(query_swerp, conn)

    # BOM 테이블과 swerp 테이블을 품목코드(Item_Code)를 기준으로 병합
    merged = orderBOM.merge(swerp, left_on='품목코드', right_on='Item_Code', how='left')

    # 필요한 열만 선택
    result = merged[['품목코드', 'Item_Name', 'Model'] + product_codes]

    # Create a dictionary to map 반제품코드 to Model
    model_mapping = dict(zip(swerp['Item_Code'], swerp['Model']))

    # Replace 반제품코드 with Model in the product_codes list
    new_product_codes = [model_mapping.get(code, code) for code in product_codes]

    # Update the column names with the new_product_codes list
    result.columns = ['품목코드', 'Item_Name', 'Model'] + new_product_codes

    print(result)
    # Save the result as an Excel file
    result.to_excel("output_2nd.xlsx", index=False)



def on_select():
    selected_codes = [listbox.item(i, 'values')[0]
                      for i in listbox.selection()]
    merge_query(selected_codes)


root = tk.Tk()
root.title("Select Product Codes and Models")

listbox = ttk.Treeview(root, selectmode="extended")
listbox["columns"] = ("Product Code", "Product Model")
listbox.column("#0", width=0, stretch=tk.NO)
listbox.column("Product Code", anchor="w", width=100)
listbox.column("Product Model", anchor="w", width=100)
listbox.heading("Product Code", text="Product Code", anchor="w")
listbox.heading("Product Model", text="Product Model", anchor="w")

product_codes_and_models = get_selected_product_codes(conn)
for _, row in product_codes_and_models.iterrows():
    listbox.insert("", "end", values=(row['반제품코드'], row['Model']))

listbox.pack(pady=10)

select_button = ttk.Button(root, text="Select", command=on_select)
select_button.pack()

root.mainloop()

conn.close()
