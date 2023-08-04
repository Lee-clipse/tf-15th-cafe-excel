import process


def get_interval_report():
    begin_date = '230726'
    end_date = '230803'
    # 수익, 비용 종합 기록 연산
    income_report = process.get_income_report(begin_date, end_date)
    outcome_report = process.get_outcome_report(begin_date, end_date)
    # 결과 엑셀 생성
    process.create_excel(income_report, outcome_report, begin_date, end_date)


if __name__ == '__main__':
    print("Hi 👋")
    get_interval_report()
