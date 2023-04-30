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

    valid_product_codes = merged.columns.intersection(product_codes)
    invalid_product_codes = set(product_codes) - set(valid_product_codes)

    if invalid_product_codes:
        print(f"Warning: The following product codes were not found: {', '.join(invalid_product_codes)}")

    # 필요한 열만 선택
    result = merged[['품목코드', 'Item_Name', 'Model'] + list(valid_product_codes)]

    print(result)
    # Save the result as an Excel file
    result.to_excel("output_2nd.xlsx", index=False)




def on_select():
    selected_codes = [listbox.item(i, 'values')[0]
                      for i in listbox.selection()]
    merge_query(selected_codes)

def search_listbox():
    search_query = search_entry.get().lower()
    listbox.delete(*listbox.get_children())

    for _, row in product_codes_and_models.iterrows():
        product_code = row['반제품코드']
        model = row['Model']

        if search_query in model.lower():
            listbox.insert("", "end", values=(product_code, model))


def on_key_release(event):
    if event.keysym == "Return":
        search_listbox()

def add_selected_items():
    selected_codes = [listbox.item(i, 'values')[1] for i in listbox.selection()]
    for code in selected_codes:
        selected_items_text.insert(tk.END, f"{code}\n")


def print_result():
    selected_codes = selected_items_text.get('1.0', tk.END).split('\n')[:-1]
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
search_label = ttk.Label(root, text="Search:")
search_label.pack(side=tk.LEFT, padx=(10, 0))

search_entry = ttk.Entry(root)
search_entry.pack(side=tk.LEFT, padx=(0, 10))
search_entry.bind("<KeyRelease>", on_key_release)

product_codes_and_models = get_selected_product_codes(conn)
for _, row in product_codes_and_models.iterrows():
    listbox.insert("", "end", values=(row['반제품코드'], row['Model']))


listbox.pack(pady=(10, 0))

add_button = ttk.Button(root, text="Add Selected", command=add_selected_items)
add_button.pack(pady=(10, 0))

selected_items_text = tk.Text(root, height=10, width=30, wrap=tk.WORD)
selected_items_text.pack(pady=(10, 0))

print_button = ttk.Button(root, text="Print Result", command=print_result)
print_button.pack(pady=(10, 0))

search_listbox()
root.mainloop()

conn.close()
