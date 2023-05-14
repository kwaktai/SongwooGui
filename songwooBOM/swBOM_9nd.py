import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pyodbc
import pandas as pd
from sqlalchemy import create_engine
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")


class FinishedProductCodeSelector:

    def __init__(self, conn_str):
        self.conn = pyodbc.connect(conn_str)
        self.finished_product_codes_and_models = self.get_finished_product_codes()

        self.root = tk.Tk()
        self.root.title("Select Finished Product Codes and Models")
        self.root.attributes('-topmost', True)
        self.root.minsize(400, 0)

        self.create_widgets()
        self.search_listbox()

    def get_semi_finished_product_codes(self, selected_codes):
        # 이 함수의 로직을 작성합니다. 예를 들어:
        semi_finished_product_codes = []
        for code in selected_codes:
            # code를 사용하여 관련 반제품 코드를 검색하고 semi_finished_product_codes에 추가합니다.
            pass
        return semi_finished_product_codes

    def get_finished_product_codes(self):
        query = """
        SELECT DISTINCT tb완제품BOMq_반제품기준.완제품코드, swerp_품목코드_쿼리.Model
        FROM tb완제품BOMq_반제품기준
        INNER JOIN swerp_품목코드_쿼리 ON tb완제품BOMq_반제품기준.완제품코드 = swerp_품목코드_쿼리.Item_Code
        """
        try:
            return pd.read_sql(query, self.conn)
        except Exception as e:
            print("Error:", e)
            return pd.DataFrame(columns=["완제품코드", "Model"])

    def get_selected_semi_finished_product_codes(self, finished_product_codes):
        if not finished_product_codes:
            return

        codes_str = ", ".join(f"'{code}'" for code in finished_product_codes)
        query = f"""
            SELECT 반제품코드, 소요량
            FROM tb완제품BOMq_반제품기준
            WHERE 완제품코드 IN ({codes_str})
            """
        return pd.read_sql(query, self.conn)

    def merge_query(self, product_codes):
        if not product_codes:
            return

        codes_str = ", ".join(f"'{code}'" for code in product_codes)

        query_semi_finished = f"""
            SELECT * FROM tb완제품BOMq_반제품기준
            WHERE 완제품코드 IN ({codes_str})
        """

        query_raw_materials = f"""
            TRANSFORM Sum(소요량)
            SELECT 품목코드
            FROM swerp_품목코드_쿼리
            WHERE 반제품코드 IN (
                SELECT 반제품코드 FROM tb완제품BOMq_반제품기준
                WHERE 완제품코드 IN ({codes_str})
            )
            GROUP BY 품목코드
            PIVOT 반제품코드;
        """

        semi_finished_product_codes = pd.read_sql(query_semi_finished, self.conn)
        raw_material_codes = pd.read_sql(query_raw_materials, self.conn)

        merged = pd.merge(raw_material_codes, semi_finished_product_codes, on='반제품코드', suffixes=('_x', '_y'))

        # Check if the columns exist before performing the operation
        if '소요량_x' in merged.columns and '소요량_y' in merged.columns:
            merged['소요량_x'] *= merged['소요량_y']
        else:
            print("Error: 소요량_x or 소요량_y column not found in the merged DataFrame")

        # Rename the columns if needed
        merged.rename(columns={'소요량_x': '소요량'}, inplace=True)

        return merged

    def on_select(self):
        selected_items = self.listbox.selection()
        selected_codes = [self.listbox.item(item)["values"][0] for item in selected_items]

        if not selected_codes:
            messagebox.showwarning("Warning", "No product code selected.")
        else:
            selected_semi_finished_product_codes = self.get_semi_finished_product_codes(selected_codes)
            if selected_semi_finished_product_codes is not None:
                self.merge_query(selected_codes)
            else:
                messagebox.showwarning("Warning", "No semi-finished product codes found.")


    def search_listbox(self):
        search_query = self.search_entry.get().lower()
        self.listbox.delete(*self.listbox.get_children())

        filtered_rows = self.finished_product_codes_and_models[self.finished_product_codes_and_models['Model'].str.lower(
        ).str.contains(search_query)]
        for _, row in filtered_rows.iterrows():
            self.listbox.insert("", "end", values=(row['완제품코드'], row['Model']))

    def on_key_release(self, event):
        if event.keysym == "Return":
            self.search_listbox()

    def create_widgets(self):
        self.listbox = ttk.Treeview(self.root, selectmode="extended")
        self.listbox["columns"] = ("Finished Product Code", "Product Model")
        self.listbox.column("#0", width=0, stretch=tk.NO)
        self.listbox.column("Finished Product Code", anchor="w", width=100)
        self.listbox.column("Product Model", anchor="w", width=280)
        self.listbox.heading("Finished Product Code", text="Finished Product Code", anchor="w")
        self.listbox.heading("Product Model", text="Product Model", anchor="w")

        self.search_label = ttk.Label(self.root, text="Search:")
        self.search_label.pack(side=tk.LEFT, padx=(10, 0))

        self.search_entry = ttk.Entry(self.root)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.search_entry.bind("<KeyRelease>", self.on_key_release)

        for _, row in self.finished_product_codes_and_models.iterrows():
            self.listbox.insert("", "end", values=(row['완제품코드'], row['Model']))

        self.listbox.pack(pady=(10, 0), padx=(0, 10))

        self.print_button = ttk.Button(
            self.root, text="Print Result", command=self.on_select)
        self.print_button.pack(pady=(10, 0))
    
    def run(self):
        self.root.mainloop()
        self.conn.close()
    

if __name__ == "__main__":
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        r"DBQ=D:\OneDrive\SongwooDB\songwoo.accdb;"
    )
    # 'engine = create_engine(f'access+pyodbc:///?odbc_connect={conn_str}')
    app = FinishedProductCodeSelector(conn_str)
    app.run()