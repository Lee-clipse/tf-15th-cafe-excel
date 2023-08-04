import utils
import pandas as pd
import openpyxl


excel = utils.get_excel()
stock_sheet = excel['재고 조사']
handmade_stock_sheet = excel['핸드메이드 재고 조사']


def get_income_report(begin_date, end_date):
    # 원하는 날짜의 데이터만 추출
    stock_report = utils.extract_interval(stock_sheet, begin_date, end_date)
    # 재고 수량 연산을 위해 데이터 가공
    stock_report = utils.accumulate_stock(stock_report)
    # 매출에 대해 합 연산
    stock_report = utils.sum_income(stock_report)

    # 핸드메이드 재고 조사 시트에 대해서 동일하게 수행

    # 원하는 날짜의 데이터만 추출
    handmade_report = utils.extract_interval(handmade_stock_sheet, begin_date, end_date)
    # 재고 수량 연산을 위해 데이터 가공
    handmade_report = utils.accumulate_handmade_stock(handmade_report)
    # 매출에 대해 합 연산
    handmade_report = utils.sum_income(handmade_report)

    # 재고 조사 시트를 병합
    report = pd.concat([stock_report, handmade_report], ignore_index=True)
    # 수량과 매출을 분리
    report = utils.divide_quantity_and_income(report)
    return report


def get_outcome_report(begin_date, end_date):
    # 원하는 날짜의 데이터만 추출
    stock_report = utils.extract_interval(stock_sheet, begin_date, end_date)
    # 입고 수량, 가격에 대해 합 연산
    stock_report = utils.accumulate_purchase_columns(stock_report)

    # 핸드메이드 재고 조사 시트에 대해서 동일하게 수행

    # 원하는 날짜의 데이터만 추출
    handmade_report = utils.extract_interval(handmade_stock_sheet, begin_date, end_date)
    # 입고 수량, 가격 대해 합 연산
    handmade_report = utils.accumulate_purchase_columns(handmade_report)

    # 재고 조사 시트를 병합
    report = pd.concat([stock_report, handmade_report], ignore_index=True)
    return report


def create_excel(income_report, outcome_report, begin_date, end_date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = begin_date + ' ~ ' + end_date + ' report'
    # 헤더 작성
    utils.append_header(ws, income_report)
    # 수익, 비용, 순이익 작성
    utils.append_income_outcome_net_profit(ws, income_report, outcome_report)
    wb.save(f"{ws.title}.xlsx")
    wb.close()
    return
