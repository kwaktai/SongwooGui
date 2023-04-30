from tkinter import Tk, Label, Button, Entry, StringVar, Text
from tkcalendar import DateEntry
from datetime import datetime
import pyodbc
from datetime import date
from dateutil.relativedelta import relativedelta

# 액세스 파일에 연결하는 함수


def connect_to_access():
    access_file = r'D:\OneDrive\SongwooDB\songwoo.accdb'
    connection_string = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={access_file};'
    )
    return pyodbc.connect(connection_string)


def format_date(date_str):
    date_obj = datetime.strptime(date_str, "%Y%m%d")
    return date_obj.strftime("%Y-%m-%d")


def _validate_date(self, date):
    try:
        datetime.strptime(date, "%Y%m%d")
        return True
    except ValueError:
        return False

# 쿼리 결과를 가져오는 함수


def get_query_result(start_date, end_date, company_keyword):
    conn = connect_to_access()
    cursor = conn.cursor()

    formatted_start_date = format_date(start_date)
    formatted_end_date = format_date(end_date)
    formatted_end_16date = start_date[:6] + "16"

    search_companies = f"""
        SELECT Company_Han, Inwhs_Com_Code
        FROM swerp_vba_Inwhs_거래처마감확인
        WHERE Company_Han LIKE '%{company_keyword}%'
        GROUP BY Company_Han, Inwhs_Com_Code
    """
    cursor.execute(search_companies)
    result_companies = cursor.fetchall()

    company_results = []

    for company in result_companies:
        company_name = company[0]
        company_code = company[1]

        print(company_name, company_code)

        query_total = f"""
            SELECT SUM(Inwhs_Amt) as total_amount
            FROM swerp_vba_Inwhs_거래처마감확인
            WHERE Metrl_Inwhs_No_Date >= #{formatted_start_date}# AND Metrl_Inwhs_No_Date <= #{formatted_end_date}# AND Inwhs_Com_Code = '{company_code}'
        """
        cursor.execute(query_total)
        result_total = cursor.fetchone()

        query_before_15th = f"""
            SELECT SUM(Inwhs_Amt) as total_amount
            FROM swerp_vba_Inwhs_거래처마감확인
            WHERE Metrl_Inwhs_No_Date >= #{formatted_start_date}# AND Metrl_Inwhs_No_Date <= #{formatted_start_date[:8]}15# AND Inwhs_Com_Code = '{company_code}'
        """
        cursor.execute(query_before_15th)
        result_before_15th = cursor.fetchone()

        company_result = (
            company_name,
            "{:,.0f}".format(result_total[0]) if result_total[0] else "0",
            "{:,.0f}".format(
                result_before_15th[0]) if result_before_15th[0] else "0",
            "{:,.0f}".format(result_total[0] - result_before_15th[0]
                             ) if result_total[0] and result_before_15th[0] else "0"
        )
        company_results.append(company_result)

    return company_results


class CustomDateEntry(DateEntry):
    def _validate_date(self, date):
        try:
            datetime.strptime(date, "%Y%m%d")
            return True
        except ValueError:
            return False


def create_gui():
    window = Tk()
    window.title("거래처 월별 마감 내역")
    window.attributes('-topmost', True)  # 창이 항상 위에 있도록 설정

    # 시작일 입력
    Label(window, text="시작일:").grid(row=0, column=0)
    today = datetime.today().strftime("%Y%m%d")  # 오늘 날짜를 가져옵니다.
    start_date = StringVar(value=today[:6] + "01")
    start_date_entry = Entry(
        window, textvariable=start_date)
    start_date_entry.grid(row=0, column=1)

    # 종료일 입력
    Label(window, text="종료일:").grid(row=1, column=0)
    end_date = StringVar(value=today)
    end_date_entry = Entry(
        window, textvariable=end_date)
    end_date_entry.grid(row=1, column=1)

    # 거래처명 입력
    Label(window, text="거래처명:").grid(row=2, column=0)
    company_name = StringVar()
    company_name_entry = Entry(window, textvariable=company_name)
    company_name_entry.grid(row=2, column=1)

    # 결과 보기 버튼
    def show_result(event=None):
        start_date_str = start_date.get()
        end_date_str = end_date.get()
        company_keyword = company_name.get()

        if company_keyword.strip():  # 거래처명이 비어있지 않은 경우에만 검색 시작
            company_results = get_query_result(
                start_date_str, end_date_str, company_keyword)

            result_text.delete(1.0, "end")
            for result in company_results:
                result_text.insert("end", f"{result[0]}\n")
                result_text.insert("end", f"총금액: {result[1]}\n")
                result_text.insert("end", f"15일 이전: {result[2]}\n")
                result_text.insert("end", f"15일 이후: {result[3]}\n")
                result_text.insert("end", "\n")

    def decrease_month():
        start_date_str = start_date.get()
        end_date_str = end_date.get()

        start_date_obj = datetime.strptime(start_date_str, "%Y%m%d")
        end_date_obj = datetime.strptime(end_date_str, "%Y%m%d")

        new_start_date_obj = start_date_obj - relativedelta(months=1)
        new_end_date_obj = end_date_obj - relativedelta(months=1)

        new_start_date = new_start_date_obj.strftime("%Y%m%d")
        new_end_date = new_end_date_obj.strftime("%Y%m%d")

        start_date.set(new_start_date)
        end_date.set(new_end_date)

    # "<" 버튼 추가
    decrease_month_button = Button(window, text="<", command=decrease_month)
    decrease_month_button.grid(row=0, column=2)

    # 결과 보기 버튼 설정
    Button(window, text="결과 보기", command=show_result).grid(
        row=3, column=0, columnspan=2)

    # 결과 출력
    result_text = Text(window, wrap="word", height=10, width=40)
    result_text.grid(row=4, column=0, columnspan=2, rowspan=4)

    # 엔터 키를 누르면 결과 보기 실행
    window.bind('<Return>', show_result)

    window.mainloop()


create_gui()
