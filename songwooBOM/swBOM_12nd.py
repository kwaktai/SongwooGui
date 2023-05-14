import tkinter as tk
from tkinter import ttk
import pyodbc
import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")


class InventoryGUI:

    def __init__(self, conn_str):
        self.conn = pyodbc.connect(conn_str)

        self.root = tk.Tk()
        self.root.title("Inventory")
        self.root.attributes('-topmost', True)
        self.root.minsize(400, 0)

        self.create_widgets()
        self.search_listbox()

    def create_widgets(self):
        self.search_label = ttk.Label(self.root, text="Search:")
        self.search_label.pack(side=tk.LEFT, padx=(10, 0))

        self.search_entry = ttk.Entry(self.root)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.search_entry.bind("<KeyRelease>", self.on_key_release)

        self.listbox = ttk.Treeview(self.root, selectmode="extended")
        self.listbox["columns"] = ("Item_Code", "Item_Name", "Model")
        self.listbox.column("#0", width=0, stretch=tk.NO)
        self.listbox.column("Item_Code", anchor="w", width=100)
        self.listbox.column("Item_Name", anchor="w", width=280)
        self.listbox.column("Model", anchor="w", width=280)
        self.listbox.heading("Item_Code", text="Item_Code", anchor="w")
        self.listbox.heading("Item_Name", text="Item_Name", anchor="w")
        self.listbox.heading("Model", text="Model", anchor="w")

        self.listbox.pack(pady=(10, 0), padx=(0, 10))

        self.print_button = ttk.Button(
            self.root, text="Print Result", command=self.print_result)
        self.print_button.pack(pady=(10, 0))

    def on_key_release(self, event):
        if event.keysym == "Return":
            self.search_listbox()

    def search_listbox(self):
        search_query = self.search_entry.get()
        if not search_query:
            self.listbox.delete(*self.listbox.get_children())
            return

        selected_items = search_query.split(',')

        result = self.merge_query(selected_items)

        self.listbox.delete(*self.listbox.get_children())
        for _, row in result.iterrows():
            self.listbox.insert("", "end", values=(row['Item_Code'], row['Item_Name'], row['Model']) + tuple(row[selected_items]))

    def merge_query(self, product_codes):
        if not product_codes:
            return

        codes_str = ", ".join(f"'{code.strip()}'" for code in product_codes)

        query_semi_finished = f"""
            SELECT * FROM tb완제품BOMq_반제품기준
            WHERE 완제품코드 IN ({codes_str})
        """

        query_raw_materials = f"""
            SELECT tb반제품BOMq.품목코드, tb완제품BOMq_반제품기준.완제품코드, Sum(tb반제품BOMq.소요량) AS 소요량
            FROM tb반제품BOMq INNER JOIN tb완제품BOMq_반제품기준 ON tb반제품BOMq.반제품코드 = tb완제품BOMq_반제품기준.반제품코드
            WHERE tb완제품BOMq_반제품기준.완제품코드 IN ({codes_str})
            GROUP BY tb반제품BOMq.품목코드, tb완제품BOMq_반제품기준.완제품코드;
        """

        semi_finished_df = pd.read_sql(query_semi_finished, self.conn)
        raw_materials_df = pd.read_sql(query_raw_materials, self.conn)

        result = pd.merge(
            semi_finished_df,
            raw_materials_df,
            left_on=["완제품코드", "반제품코드"],
            right_on=["완제품코드", "품목코드"],
            how="left"
        ).drop(columns=["품목코드"])
        
        result = result.rename(columns={
            "완제품코드": "Item_Code", 
            "완제품명": "Item_Name", 
            "규격": "Model", 
            "소요량": "소요량"
        })

        return result
    def print_result(self):
        selected_item = self.listbox.focus()
        if not selected_item:
            print("No item selected")
            return

        item_values = self.listbox.item(selected_item)["values"]
        if not item_values:
            print("No item selected")
            return

        print(f"Item Code: {item_values[0]}, Item Name: {item_values[1]}, Model: {item_values[2]}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    # conn_string = "<your_connection_string>"
    conn_string = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        r"DBQ=D:\OneDrive\SongwooDB\songwoo.accdb;"
    )
    app = InventoryGUI(conn_string)
    app.run()
# 해당 코드는 완제품코드로 반제품코드를 검색하여, BOM 를 구성하는것입니다.
# 하지만, 완제품코드로 표현되어 규격으로 만들지 못했습니다.
