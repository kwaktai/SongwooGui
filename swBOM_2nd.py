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
    query = "SELECT [반제품코드] FROM [tb반제품BOMq] GROUP BY [반제품코드];"
    return pd.read_sql(query, conn)['반제품코드'].tolist()

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
    print("codes_str: " + codes_str)
    print("product_codes: " + product_codes)





    # Save the result as an Excel file
    # result.to_excel("output_2nd.xlsx", index=False)



def on_select():
    selected_codes = [listbox.item(i, 'values')[0] for i in listbox.selection()]
    create_pivot_table(conn, selected_codes)

root = tk.Tk()
root.title("Select Product Codes")

listbox = ttk.Treeview(root, selectmode="extended")
listbox["columns"] = ("Product Code",)
listbox.column("#0", width=0, stretch=tk.NO)
listbox.column("Product Code", anchor="w", width=100)
listbox.heading("Product Code", text="Product Code", anchor="w")

product_codes = get_selected_product_codes(conn)
for code in product_codes:
    listbox.insert("", "end", values=(code,))

listbox.pack(pady=10)

select_button = ttk.Button(root, text="Select", command=on_select)
select_button.pack()

root.mainloop()

conn.close()
