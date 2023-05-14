import pyodbc
import pandas as pd
from collections import deque
import pandas as pd
from openpyxl import Workbook

# DB 연결 설정
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=D:\OneDrive\SongwooDB\songwoo.accdb;'
)
cnxn = pyodbc.connect(conn_str)

# 테이블과 쿼리 로드
basic_stock = pd.read_sql_query('SELECT * FROM Basic_Stock', cnxn)
inwhs = pd.read_sql_query('SELECT * FROM swerp_vba_Inwhs ORDER BY Metrl_Inwhs_No_Date', cnxn)
outwhs = pd.read_sql_query('SELECT * FROM swerp_vba_outwhs ORDER BY Metrl_Outwhs_No_Date', cnxn)

# 기초재고 설정
stock = {row.Item_Code: deque([(row.Basic_Qty, row.basic_price)]) for idx, row in basic_stock.iterrows()}
# last_in_price 딕셔너리 생성
last_in_price = {}

# 입고내역 처리
for idx, row in inwhs.iterrows():
    if row.Item_Code not in stock:
        stock[row.Item_Code] = deque()
    stock[row.Item_Code].append((row.Inwhs_Qty, row.Inwhs_Price))
    last_in_price[row.Item_Code] = row.Inwhs_Price  # 마지막 입고 가격 업데이트


# 출고내역 처리
for idx, row in outwhs.iterrows():
    if row.Item_Code in stock:
        out_qty = row.Outwhs_Qty
        while out_qty > 0 and stock[row.Item_Code]:
            qty, price = stock[row.Item_Code][0]
            if qty <= out_qty:
                out_qty -= qty
                stock[row.Item_Code].popleft()
            else:
                stock[row.Item_Code][0] = (qty - out_qty, price)
                out_qty = 0


# 최종 재고와 단가 계산
final_stock = {item: sum(qty for qty, price in deq) for item, deq in stock.items()}
final_price = {item: round(sum(qty*price for qty, price in deq)/final_stock[item], 2) if final_stock[item] > 0 else last_in_price.get(item, 0) for item, deq in stock.items()}
# pandas와 openpyxl 라이브러리를 import합니다. (openpyxl은 pandas가 Excel 파일을 쓸 때 필요한 라이브러리입니다.)

# 딕셔너리를 DataFrame으로 변환합니다.
df_stock = pd.DataFrame(list(final_stock.items()), columns=['Item_Code', 'Final_Stock'])
df_price = pd.DataFrame(list(final_price.items()), columns=['Item_Code', 'Final_Price'])

# 두 데이터프레임을 Item_Code를 기준으로 합칩니다.
df_final = pd.merge(df_stock, df_price, on='Item_Code')

# 데이터프레임을 출력합니다.
print(df_final)

# 데이터프레임을 엑셀 파일로 저장합니다.
df_final.to_excel('songwooFIFO\output_FIFO.xlsx', index=False)


# 외주업체 기초재고도 전부 넣어야 완성 되겠다.
