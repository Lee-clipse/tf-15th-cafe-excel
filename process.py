import openpyxl
import pandas as pd

import utils

sheet_list = utils.get_sheet_list()
receive_log = utils.get_receive_log(sheet_list)
stock_survey_sheet_name_list = utils.get_stock_survey_sheet_name_list(sheet_list)


# 호출: main - get_interval_report()
# 동작: 시트들을 순회하며 특정 날짜 구간 안에 속하는 시트의 데이터를 일렬로 반환
# 반환: 종류 / 상품명 / 출고 수량 / 판매 수입 / 일자
def get_sales_straight_log(begin_date, end_date):
    log = pd.DataFrame([])
    for sheet_name in stock_survey_sheet_name_list:
        date = sheet_name[:6]
        if not utils.is_include_date(date, begin_date, end_date):
            continue
        stock_survey = utils.set_product_kind_valid(sheet_list[sheet_name])
        log = pd.concat([log, utils.get_sales_log(stock_survey, date)], axis=0, ignore_index=True)
    return log


# 호출: main - get_weekday_report()
# 동작: 시트들을 순회하며 특정 요일에 해당하는 시트의 데이터를 일렬로 반환
# 반환: 종류 / 상품명 / 출고 수량 / 판매 수입 / 일자
def get_sales_straight_log_for_weekday(weekday):
    log = pd.DataFrame([])
    for sheet_name in stock_survey_sheet_name_list:
        if not weekday == sheet_name[7]:
            continue
        date = sheet_name[:6]
        stock_survey = utils.set_product_kind_valid(sheet_list[sheet_name])
        log = pd.concat([log, utils.get_sales_log(stock_survey, date)], axis=0, ignore_index=True)
    return log


# 호출: main - get_interval_report(), get_weekday_report()
# 동작: 매출 리포트를 작성
def get_sales_report(sales_straight_log):
    sales_report = sales_straight_log\
        .groupby(['일자', '종류'])\
        .agg({'출고 수량': 'sum', '판매 수입': 'sum'})
    # 상품 종류 컨벤션을 맞추기 위함
    date_list = sorted(list(set([row.name[0] for index, row in sales_report.iterrows()])))
    sales_report = utils.append_kind_polyfill(sales_report, date_list)
    sales_report = sales_report\
        .groupby(['일자', '종류'])\
        .agg({'출고 수량': 'sum', '판매 수입': 'sum'})\
        .sort_values(by=['일자', '종류'], ascending=[True, True])
    return sales_report


# 호출: main - get_interval_report()
# 동작: 입고 Log를 순회하며 특정 날짜 구간 안에 속하는 row의 데이터를 반환
# 반환: 종류 / 상품명 / 입고 수량 / 구매 비용 / 일자
def get_expenses_straight_log(begin_date, end_date):
    log = pd.DataFrame([])
    for index, row in receive_log.iterrows():
        row_date = str(row['일자'])
        if not utils.is_include_date(row_date, begin_date, end_date):
            continue
        log = pd.concat([log, utils.get_expenses_log(row)], axis=0, ignore_index=True)
    return pd.DataFrame(log)


# 호출: main - get_interval_report(), get_weekday_report()
# 동작: 매입 리포트를 작성
def get_expenses_report(expenses_straight_log):
    expenses_report = expenses_straight_log\
        .groupby(['일자', '종류'])\
        .agg({'입고 수량': 'sum', '구매 비용': 'sum'})\
        .sort_values(by=['일자', '종류'], ascending=[True, True])
    # 상품 종류 컨벤션을 맞추기 위함
    date_list = sorted(list(set([row.name[0] for index, row in expenses_report.iterrows()])))
    expenses_report = utils.append_kind_polyfill(expenses_report, date_list)
    expenses_report = expenses_report\
        .groupby(['일자', '종류'])\
        .agg({'입고 수량': 'sum', '구매 비용': 'sum'})\
        .sort_values(by=['일자', '종류'], ascending=[True, True])
    return expenses_report


# 호출: main - get_interval_report(), get_weekday_report()
# 동작: 랭킹 기록을 위해 모든 날짜의 상품 기록을 일렬로 반환
def get_product_sales_ranking(sales_straight_log):
    return sales_straight_log\
        .groupby(['종류', '상품명'])\
        .agg({'출고 수량': 'sum', '판매 수입': 'sum'})\
        .sort_values(by='판매 수입', ascending=False)


# 호출: main - get_interval_report(), get_weekday_report()
# 동작: 종류 별 상품의 매출 데이터를 추출하기 위해 가공
def get_product_sales_report(sales_straight_log):
    return sales_straight_log\
        .groupby(['종류', '상품명'])\
        .agg({'출고 수량': 'sum', '판매 수입': 'sum'})\
        .sort_values(by=['종류', '판매 수입'], ascending=[True, False])


# 호출: main - get_interval_report()
# 동작: 종류 별 상품의 매입 데이터를 추출하기 위해 가공
def get_product_expense_report(expenses_straight_log):
    return expenses_straight_log\
        .groupby(['종류', '상품명'])\
        .agg({'입고 수량': 'sum', '구매 비용': 'sum'})\
        .sort_values(by=['종류', '구매 비용'], ascending=[True, False])


# 호출: main - get_interval_report()
# 동작: 종합 매입 매출 데이터 엑셀을 작성
def create_sales_expenses_overall_report(
        begin_date, end_date, sales_report, expenses_report, product_sales_ranking, sales_straight_log, expenses_straight_log):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = begin_date + ' ~ ' + end_date + ' overall report'

    # 매출 항목을 작성
    ws.append(['매출'])
    index = 2
    utils.append_product_kind_heads(ws, index)
    utils.append_heads_columns(ws, index+1, '출고 수량', '판매 수입')
    utils.insert_sales_data(ws, sales_report)
    utils.insert_sales_sum_data(ws, sales_straight_log)

    # 현재 작성중인 row의 index를 계산
    sales_row_count = len(sales_report.index.get_level_values('일자').unique())
    index = 7 + sales_row_count
    ws.append([])

    # 매입 항목을 작성
    ws.append(['지출'])
    utils.append_product_kind_heads(ws, index)
    utils.append_heads_columns(ws, index+1, '입고 수량', '구매 비용')
    utils.insert_expenses_data(ws, expenses_report)
    utils.insert_expenses_sum_data(ws, expenses_straight_log)

    # 현재 작성중인 row의 index를 계산
    expenses_row_count = len(expenses_report.index.get_level_values('일자').unique())
    ws.append([])

    # 매출, 매입을 종합한 순이익을 작성
    ws.append(['순이익'])
    utils.append_product_kind_short_heads(ws)
    utils.insert_net_profit_row(ws, sales_row_count, expenses_row_count)

    # 상품별 매출 랭킹을 작성
    ws.append([])
    ws.append(['판매 순위'])
    utils.append_sales_ranking_heads(ws)
    utils.insert_sales_ranking_data(ws, product_sales_ranking)

    # 엑셀 파일을 현재 위치에 저장
    wb.save(f"{ws.title}.xlsx")
    wb.close()
    print(f"스냅샷이 {ws.title}.xlsx 파일로 저장되었습니다.")
    return


# 호출: main - get_interval_report()
# 동작: 종류별 상품 매입 매출 데이터 엑셀을 작성
def create_sales_expenses_product_report(begin_date, end_date, product_sales_report, product_expense_report):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = begin_date + ' ~ ' + end_date + ' product report'

    # 매출 항목을 작성
    ws.append(['매출'])
    utils.insert_sales_ranking_data_at_product(ws, product_sales_report)

    # 매입 항목을 작성
    ws.append([])
    ws.append(['지출'])
    utils.insert_expenses_ranking_data_at_product(ws, product_expense_report)

    # 엑셀 파일을 현재 위치에 저장
    wb.save(f"{ws.title}.xlsx")
    wb.close()
    print(f"스냅샷이 {ws.title}.xlsx 파일로 저장되었습니다.")
    return


# 호출: main - get_weekday_report()
# 동작: 특정 요일에 대한 종합 매입 매출 데이터 엑셀을 작성
def create_weekday_overall_report(weekday, sales_report, product_sales_ranking, sales_straight_log):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = weekday + '요일 overall report'

    # 매출 항목을 작성
    ws.append(['매출'])
    index = 2
    utils.append_product_kind_heads(ws, index)
    utils.append_heads_columns(ws, index+1, '출고 수량', '판매 수입')
    utils.insert_sales_data(ws, sales_report)
    utils.insert_sales_sum_data(ws, sales_straight_log)

    # 상품별 매출 랭킹을 작성
    ws.append([])
    ws.append(['판매 순위'])
    utils.append_sales_ranking_heads(ws)
    utils.insert_sales_ranking_data(ws, product_sales_ranking)

    # 엑셀 파일을 현재 위치에 저장
    wb.save(f"{ws.title}.xlsx")
    wb.close()
    print(f"스냅샷이 {ws.title}.xlsx 파일로 저장되었습니다.")
    return


# 호출: main - get_weekday_report()
# 동작: 특정 요일에 대한 종류별 상품 매입 매출 데이터 엑셀을 작성
def create_weekday_product_report(weekday, product_sales_report):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = weekday + '요일 product report'

    # 매출 항목을 작성
    ws.append(['매출'])
    utils.insert_sales_ranking_data_at_product(ws, product_sales_report)

    # 엑셀 파일을 현재 위치에 저장
    wb.save(f"{ws.title}.xlsx")
    wb.close()
    print(f"스냅샷이 {ws.title}.xlsx 파일로 저장되었습니다.")
    return
