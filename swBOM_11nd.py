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
            SELECT tb반제품BOMq.품목코드, tb완제품BOMq_반제품기준.완제품코드, Sum(tb반제품BOMq.소요량) AS 소요량
            FROM tb반제품BOMq INNER JOIN tb완제품BOMq_반제품기준 ON tb반제품BOMq.반제품코드 = tb완제품BOMq_반제품기준.반제품코드
            WHERE tb완제품BOMq_반제품기준.완제품코드 IN ({codes_str})
            GROUP BY tb반제품BOMq.품목코드, tb완제품BOMq_반제품기준.완제품코드;
        """

        query_item_info = f"""
            SELECT swerp_품목코드_쿼리.Item_Code, swerp_품목코드_쿼리.Item_Name, swerp_품목코드_쿼리.Model
            FROM swerp_품목코드_쿼리;
        """


        semi_finished_product_codes =  pd.read_sql(query_semi_finished, self.conn)
        raw_material_codes = pd.read_sql(query_raw_materials, self.conn)
        item_info = pd.read_sql(query_item_info, self.conn)

        pivot_df = raw_material_codes.pivot_table(index='품목코드', columns='완제품코드', values='소요량')
        pivot_df = pivot_df.reset_index().rename(columns={'품목코드': 'Item_Code'})
        result = item_info.merge(pivot_df, on='Item_Code', how='right')

        return result

inventory = Inventory("D:/OneDrive/SongwooDB/songwoo.accdb")

selected_items = ["1011031", "1011029"]

result = inventory.merge_query(selected_items)

print(result)
