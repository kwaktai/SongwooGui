import pandas as pd
import pyodbc

class Inventory:
    def __init__(self, db_path):
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={db_path};'
        )
        self.conn = pyodbc.connect(conn_str)

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
            FROM tb반제품BOMq
            WHERE 반제품코드 IN (
                SELECT 반제품코드 FROM tb완제품BOMq_반제품기준
                WHERE 완제품코드 IN ({codes_str})
            )
            GROUP BY 품목코드
            PIVOT 반제품코드;
        """

        semi_finished_product_codes = pd.read_sql(query_semi_finished, self.conn)
        raw_material_codes = pd.read_sql(query_raw_materials, self.conn)

        return raw_material_codes

inventory = Inventory("D:/OneDrive/SongwooDB/songwoo.accdb")

selected_items = ["1011031","1011029"]
result = inventory.merge_query(selected_items)

print(result)
