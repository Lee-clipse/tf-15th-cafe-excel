import process


def get_interval_report():
    # 날짜 입력
    # begin_date = input('시작 날짜를 입력하세요 (ex: 230630) >> ')
    # end_date = input('종료 날짜를 입력하세요 (ex: 230630) >> ')
    begin_date = '230630'
    end_date = '230715'
    # 매출 종합 계산
    sales_straight_log = process.get_sales_straight_log(begin_date, end_date)
    sales_report = process.get_sales_report(sales_straight_log)
    # 지출 종합 계산
    expenses_straight_log = process.get_expenses_straight_log(begin_date, end_date)
    expenses_report = process.get_expenses_report(expenses_straight_log)
    # 판매 순위 계산
    product_sales_ranking = process.get_product_sales_ranking(sales_straight_log)
    # 종류 내 상품별 매출 종합 계산
    product_sales_report = process.get_product_sales_report(sales_straight_log)
    # 종류 내 상품별 지출 종합 계산
    product_expense_report = process.get_product_expense_report(expenses_straight_log)
    # 엑셀 생성
    process.create_sales_expenses_overall_report(
        begin_date, end_date, sales_report, expenses_report, product_sales_ranking, sales_straight_log, expenses_straight_log)
    process.create_sales_expenses_product_report(begin_date, end_date, product_sales_report, product_expense_report)
    return


def get_weekday_report():
    weekday = input('요일을 입력하세요 (ex: 수, 금, 일) >> ')
    # 매출 종합 계산
    sales_straight_log = process.get_sales_straight_log_for_weekday(weekday)
    sales_report = process.get_sales_report(sales_straight_log)
    # 판매 순위 계산
    product_sales_ranking = process.get_product_sales_ranking(sales_straight_log)
    # 종류 내 상품별 매출 종합 계산
    product_sales_report = process.get_product_sales_report(sales_straight_log)
    # 엑셀 생성
    process.create_weekday_overall_report(weekday, sales_report, product_sales_ranking, sales_straight_log)
    process.create_weekday_product_report(weekday, product_sales_report)
    return


if __name__ == '__main__':
    print('👋')
    # 특정 구간 내 데이터를 종합하여 가공한 엑셀 파일을 생성
    get_interval_report()

    # 특정 요일의 모든 데이터를 종합하여 가공한 엑셀 파일을 생성
    # get_weekday_report()
