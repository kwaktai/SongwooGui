import tkinter as tk
from tkinter import ttk
import pyodbc
import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")


class ProductCodeSelector:

    def __init__(self, conn_str):
        self.conn = pyodbc.connect(conn_str)
        self.product_codes_and_models = self.get_selected_product_codes()

        self.root = tk.Tk()
        self.root.title("Select Product Codes and Models-8nd")
        self.root.attributes('-topmost', True)
        self.root.minsize(400, 0)

        self.create_widgets()
        self.search_listbox()

    def get_selected_product_codes(self):
        query = """
        SELECT tb반제품BOMq.반제품코드, swerp_품목코드_쿼리.Model
        FROM tb반제품BOMq
        INNER JOIN swerp_품목코드_쿼리 ON tb반제품BOMq.반제품코드 = swerp_품목코드_쿼리.Item_Code
        GROUP BY tb반제품BOMq.반제품코드, swerp_품목코드_쿼리.Model;
        """
        return pd.read_sql(query, self.conn)

    # Add other methods here like 'create_pivot_table', 'merge_query',
    # class ProductCodeSelector:

    def get_selected_product_codes(self):
        query = """
        SELECT tb반제품BOMq.반제품코드, swerp_품목코드_쿼리.Model
        FROM tb반제품BOMq
        INNER JOIN swerp_품목코드_쿼리 ON tb반제품BOMq.반제품코드 = swerp_품목코드_쿼리.Item_Code
        GROUP BY tb반제품BOMq.반제품코드, swerp_품목코드_쿼리.Model;
        """
        return pd.read_sql(query, self.conn)

    def create_pivot_table(self, product_codes):
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

        result = pd.read_sql(query, self.conn)

        # Fill missing values with 0
        result = result.fillna(0)

        return result

    def merge_query(self, product_codes):
        orderBOM = self.create_pivot_table(product_codes)

        query_swerp = "SELECT Item_Code, Item_Name, Model FROM swerp_품목코드_쿼리;"
        swerp = pd.read_sql(query_swerp, self.conn)

        # BOM 테이블과 swerp 테이블을 품목코드(Item_Code)를 기준으로 병합
        merged = orderBOM.merge(swerp, left_on='품목코드',
                                right_on='Item_Code', how='left')

        valid_product_codes = merged.columns.intersection(product_codes)
        invalid_product_codes = set(product_codes) - set(valid_product_codes)

        if invalid_product_codes:
            print(
                f"Warning: The following product codes were not found: {', '.join(invalid_product_codes)}")

        # 선택한 제품 코드를 모델 이름으로 변환
        product_models = self. product_codes_and_models.set_index('반제품코드')[
            'Model'].to_dict()
        model_names = [product_models.get(code, code)
                       for code in valid_product_codes]

        # 결과 데이터 프레임의 열 이름을 적절한 모델 이름으로 바꾸기
        merged = merged.rename(columns=dict(
            zip(valid_product_codes, model_names)))

        # 필요한 열만 선택
        result = merged[['품목코드', 'Item_Name', 'Model'] + model_names]

        print(result)
        # Save the result as an Excel file
        result.to_excel("output_2nd.xlsx", index=False)

    def on_select(self):
        selected_codes = [self.listbox.item(i, 'values')[0]
                          for i in self.listbox.selection()]
        self.merge_query(selected_codes)

    def search_listbox(self):
        search_query = self.search_entry.get().lower()
        self.listbox.delete(*self.listbox.get_children())

        filtered_rows = self.product_codes_and_models[self.product_codes_and_models['Model'].str.lower(
        ).str.contains(search_query)]
        for _, row in filtered_rows.iterrows():
            self.listbox.insert("", "end", values=(row['반제품코드'], row['Model']))

    def on_key_release(self, event):
        if event.keysym == "Return":
            self.search_listbox()

    def add_selected_items(self):
        selected_items = [self.listbox.item(i, 'values')
                          for i in self.listbox.selection()]
        for item in selected_items:
            self.selected_items_text.insert(tk.END, f"{item[0]} - {item[1]}\n")

    def print_result(self):
        selected_codes = [item.split(
            " - ")[0] for item in self.selected_items_text.get('1.0', tk.END).split('\n')[:-1]]
        self.merge_query(selected_codes)

    def create_widgets(self):
        self.listbox = ttk.Treeview(self.root, selectmode="extended")
        self.listbox["columns"] = ("Product Code", "Product Model")
        self.listbox.column("#0", width=0, stretch=tk.NO)
        self.listbox.column("Product Code", anchor="w", width=100)
        self.listbox.column("Product Model", anchor="w", width=280)
        self.listbox.heading("Product Code", text="Product Code", anchor="w")
        self.listbox.heading("Product Model", text="Product Model", anchor="w")

        self.search_label = ttk.Label(self.root, text="Search:")
        self.search_label.pack(side=tk.LEFT, padx=(10, 0))

        self.search_entry = ttk.Entry(self.root)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.search_entry.bind("<KeyRelease>", self.on_key_release)

        for _, row in self.product_codes_and_models.iterrows():
            self.listbox.insert("", "end", values=(row['반제품코드'], row['Model']))

        self.listbox.pack(pady=(10, 0), padx=(0, 10))

        self.add_button = ttk.Button(
            self.root, text="Add Selected", command=self.add_selected_items)
        self.add_button.pack(pady=(10, 0))

        self.selected_items_text = tk.Text(
            self.root, height=10, width=54, wrap=tk.WORD)
        self.selected_items_text.pack(pady=(10, 0), padx=(0, 10))

        self.print_button = ttk.Button(
            self.root, text="Print Result", command=self.print_result)
        self.print_button.pack(pady=(10, 0))

    def run(self):
        self.root.mainloop()
        self.conn.close()


if __name__ == "__main__":
    conn_str = (
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        r"DBQ=D:\OneDrive\SongwooDB\songwoo.accdb;"
    )
    app = ProductCodeSelector(conn_str)
    app.run()
