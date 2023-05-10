import tkinter as tk
from tkinter import Listbox, Button, Scrollbar, MULTIPLE
from tkinter import messagebox
from pandastable import Table
import pandas as pd
import pyodbc
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")


class Inventory:
    def __init__(self, db_path):
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={db_path};'
        )
        self.conn = pyodbc.connect(conn_str)

    def load_item_info(self):
        query_item_info = f"""
            SELECT swerp_품목코드_쿼리.Item_Code, swerp_품목코드_쿼리.Item_Name, swerp_품목코드_쿼리.Model
            FROM swerp_품목코드_쿼리;
        """
        return pd.read_sql(query_item_info, self.conn)
    
    def matching_item_info(self):
        query = f"""
            SELECT tb완제품BOMq_반제품기준.완제품코드, swerp_품목코드_쿼리.Model
            FROM tb완제품BOMq_반제품기준
            INNER JOIN swerp_품목코드_쿼리 ON tb완제품BOMq_반제품기준.완제품코드 = swerp_품목코드_쿼리.Item_Code
            GROUP BY tb완제품BOMq_반제품기준.완제품코드, swerp_품목코드_쿼리.Model;
        """

        return pd.read_sql(query, self.conn)
    
    def merge_query(self, product_codes):
        if not product_codes:
            return

        print(f"product_codes : {product_codes}")
        codes_str = ", ".join(f"'{code}'" for code in product_codes)
        print(f"codes_str : {codes_str}")

        # codes_str = "'1013070', '1013071', '1011029'"
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

        semi_finished_product_codes =  pd.read_sql(query_semi_finished, self.conn)
        raw_material_codes = pd.read_sql(query_raw_materials, self.conn)
        item_info = self.load_item_info()

        pivot_df = raw_material_codes.pivot_table(index='품목코드', columns='완제품코드', values='소요량')
        pivot_df = pivot_df.reset_index().rename(columns={'품목코드': 'Item_Code'})

        # Update column names
        new_column_names = {code: item_info[item_info['Item_Code'] == code]['Model'].values[0] if item_info[item_info['Item_Code'] == code]['Model'].values.size > 0 else code for code in pivot_df.columns if code != 'Item_Code'}
        pivot_df = pivot_df.rename(columns=new_column_names)

        result = item_info.merge(pivot_df, on='Item_Code', how='right')
        print(f"result : {result}")
        return result

class InventoryGui(tk.Tk):
    def __init__(self, inventory):
        super().__init__()

        self.inventory = inventory

        self.title("Inventory Management")

        # Create Entry for Search
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(self, textvariable=self.search_var)
        self.search_entry.bind('<Return>', self.search_items)  # bind the Enter key to the search_items method
        self.search_entry.grid(row=0, column=0, sticky='w')

        self.listbox1 = Listbox(self, selectmode=MULTIPLE, width=70, height=50)
        self.listbox1.bind('<Double-Button-1>', self.move_item)  # bind double-click to move_item method
        self.listbox1.grid(row=1, column=0)

        self.listbox2 = Listbox(self, selectmode=MULTIPLE, width=70, height=50)
        self.listbox2.bind('<Double-Button-1>', self.delete_item)  # bind double-click to delete_item method
        self.listbox2.grid(row=1, column=1)

        # Create clear button
        self.clear_btn = Button(self, text="Clear All", command=self.clear_all)
        self.clear_btn.grid(row=0, column=1)

        self.btn_select = tk.Button(
            self, text="Select", command=self.select_items)
        self.btn_select.grid(row=2, column=0)

        self.btn_result = tk.Button(
            self, text="Result", command=self.show_result)
        self.btn_result.grid(row=2, column=1)

        self.load_items()

    def load_items(self):
        # Load items from the inventory
        item_info = self.inventory.matching_item_info()
        for index, row in item_info.iterrows():
            self.listbox1.insert(
                tk.END, f"{row['완제품코드']} - {row['Model']}")

    def select_items(self):
        # Move selected items from the first box to the second box
        selected_indices = self.listbox1.curselection()
        for i in selected_indices:
            self.listbox2.insert(tk.END, self.listbox1.get(i))
        self.listbox1.delete(selected_indices)
        for index in selected_indices:
            if index < self.listbox1.size():
                self.listbox1.delete(index)
            else:
                print(f"Index {index} is out of range")

    def show_result(self):
        # Show the result for the selected items
        selected_items = [self.get_item_code(
            item) for item in self.listbox2.get(0, tk.END)]
        result = self.inventory.merge_query(selected_items)

        # Show result in a new window
        result_window = tk.Toplevel(self)
        result_window.title("Result")
        frame = tk.Frame(result_window)
        frame.pack(fill='both', expand=True)
        pt = Table(frame, dataframe=result,
                   showtoolbar=True, showstatusbar=True)
        pt.show()

        # Save result to excel
        result.to_excel(
            "D:\TaiCloud\Documents\Documents\Project\SongwooGui\songwoo.xlsx")
    
    def search_items(self, event):
        # Get the current text from the search entry
        search_term = self.search_entry.get().lower()
        self.listbox1.delete(0, tk.END)
        if not search_term:
            self.load_items()
        else:
            # Load items from the inventory
            item_info = self.inventory.matching_item_info()
            for index, row in item_info.iterrows():
                item_string = f"{row['완제품코드']} - {row['Model']}"
                if search_term in item_string.lower():
                    self.listbox1.insert(tk.END, item_string)
    
    def clear_all(self):
        self.listbox2.delete(0, tk.END)

    def move_item(self, event):
        # Move selected item from the first box to the second box
        selected = self.listbox1.curselection()
        if selected:
            self.listbox2.insert(tk.END, self.listbox1.get(selected))
            self.listbox1.delete(selected)

    def delete_item(self, event):
        # Delete selected item from the second box
        selected = self.listbox2.curselection()
        if selected:
            self.listbox2.delete(selected)

    @staticmethod
    def get_item_code(item):
        # Extract the item code from the string
        return item.split(" - ")[0]


# Create the inventory
inventory = Inventory("D:/OneDrive/SongwooDB/songwoo.accdb")

# Create and start the GUI
gui = InventoryGui(inventory)
gui.mainloop()

# 검색 기능 부족으로 추가 작업
# 1. 검색 기능 추가
# 2. 검색 기능을 통해 선택된 제품만 리스트에 추가
# 3. 선택된 제품만 리스트에 추가