from datetime import datetime
import pandas as pd


# 호출: process - global
# 동작: 엑셀 반환
def get_excel():
    return pd.read_excel(
        '카페 재고 조사표.xlsx',
        sheet_name=None,
        engine='openpyxl',
    )


# 호출: process - get_income_report()
# 동작: 사용자가 입력한 일자 구간에 맞도록 엑셀을 추출하여 반환
def extract_interval(stock_sheet, begin_date, end_date):
    delete_column_names = []
    # 포함되는 날짜의 column만 추출
    for column_name in stock_sheet.columns[3:]:
        date = column_name[:6]
        if not is_include_date(date, begin_date, end_date):
            delete_column_names.append(column_name)
    return stock_sheet.drop(delete_column_names, axis=1)


# 호출: extract_interval()
# 동작: 특정 구간 내에 포함되는 일자인지 검사
def is_include_date(date, begin_date, end_date):
    date_obj = datetime.strptime(date, "%y%m%d")
    begin_date_obj = datetime.strptime(begin_date, "%y%m%d")
    end_date_obj = datetime.strptime(end_date, "%y%m%d")
    return begin_date_obj <= date_obj <= end_date_obj


# 호출: process - get_income_report()
# 동작: 재고 조사 결과 수량만을 누적, 추출해서 반환
def accumulate_stock(report):
    # 입고 수량을 누적해서 재고 수량으로 반환
    stock_columns = get_stock_columns(report).astype(int)
    purchase_columns = get_purchase_columns(report).astype(int)
    stock_columns = stock_columns.add(purchase_columns, fill_value=0).fillna(0).astype(int)
    # 원래 양식대로 df 구성
    header_columns = get_header_columns(report)
    return pd.concat([header_columns, stock_columns], axis=1)


# 호출: process - get_income_report()
# 동작: 핸드메이드 재고 조사 시트의 판매 수량을 누적, 추출해서 반환
def accumulate_handmade_stock(report):
    header_columns = get_header_columns(report)
    stock_columns = get_stock_columns(report).astype(int)
    return pd.concat([header_columns, stock_columns], axis=1)


# 호출: accumulate_stock(), sum_income()
# 동작: 헤더 column만 추출해서 반환
def get_header_columns(report):
    header_columns = report.iloc[:, [0, 1, 2]]
    return header_columns


# 호출: accumulate_stock(), sum_income()
# 동작: 재고 조사 column만 추출해서 반환
def get_stock_columns(report):
    stock_columns = report.filter(like='재', axis=1)
    return stock_columns


# 호출: accumulate_stock()
# 동작: 입고 column을 추출하여 입고 수량만 값으로 반환
def get_purchase_columns(report):
    purchase_columns = report\
        .filter(like='입', axis=1)\
        .apply(lambda col: col.apply(lambda val: val.split('/')[0]))
    # add 함수를 사용하기 위해 column index를 변경
    purchase_columns.columns = purchase_columns.columns.str.replace('입', '재')
    return purchase_columns


# 호출: process - get_income_report()
# 동작: 재고에 대해 매출 합 연산을 수행하여 반환
def sum_income(report):
    header_columns = get_header_columns(report)
    # 재고 수량 합치기
    sum_columns = pd.DataFrame([])
    stock_columns = get_stock_columns(report)
    sum_columns['합계'] = stock_columns.sum(axis=1)
    accumulate_columns = pd.concat([header_columns, sum_columns], axis=1)
    # 재고 수량 * 판매가 => 매출
    accumulate_columns['합계'] = accumulate_columns\
        .apply(lambda col: f"{int(col['합계'])}/{col['판매가'] * int(col['합계'])}", axis=1)
    return accumulate_columns


# 호출: process - get_income_report()
# 동작: 물품의 판매 수량, 수익을 분리하여 반환
def divide_quantity_and_income(report):
    report[['수량', '수익']] = report['합계'].str.split('/', expand=True)
    return report.drop(columns='합계')


# 호출: process - get_outcome_report()
# 동작: 입고 column만 추출하여 반환
def accumulate_purchase_columns(report):
    purchase_columns = report.filter(like='입', axis=1)
    # 수량, 매출을 분리하여 합 연산
    quantity_columns = purchase_columns\
        .apply(lambda col: col.apply(lambda val: val.split('/')[0])) \
        .astype(int)\
        .sum(axis=1)
    price_columns = purchase_columns\
        .apply(lambda col: col.apply(lambda val: val.split('/')[1])) \
        .astype(int)\
        .sum(axis=1)
    # 원래 양식대로 df 구성
    header_columns = get_header_columns(report)
    result_columns = pd.DataFrame({'수량': quantity_columns, '비용': price_columns})
    return pd.concat([header_columns, result_columns], axis=1)

