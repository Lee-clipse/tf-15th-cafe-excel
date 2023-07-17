import pandas as pd
from datetime import datetime
import re


product_kind_list = ['간식', '라면, 인스턴트', '생필품', '아이스크림', '음료', '일회용품', '친환경제품', '타먹는 차']


# 호출: process - 전역
# 동작: 엑셀 호출
def get_sheet_list():
    return pd.read_excel(
        '교회 카페 관리.xlsx',
        sheet_name=None,
        engine='openpyxl',
        header=3,
        index_col=None,
        usecols="B:H"
    )


# 호출: process - 전역
# 동작: 엑셀 내 `입고 Log` 시트를 반환
def get_receive_log(sheet_list):
    return sheet_list['입고 Log']


# 호출: process - 전역
# 동작: 엑셀 내 `재고 조사` 시트들을 반환
def get_stock_survey_sheet_name_list(sheet_list):
    stock_survey_sheet_name_list = []
    for sheet_name in sheet_list:
        date_pattern = r'^\d{2}(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])'
        if re.match(date_pattern, sheet_name):
            stock_survey_sheet_name_list.append(sheet_name)
    return stock_survey_sheet_name_list


# 호출: process - straight_log 동작 함수들
# 동작: 특정 구간 내에 포함되는 일자인지 datetime 객체로 변환하여 반환
# 반환: boolean
def is_include_date(date, begin_date, end_date):
    date_obj = datetime.strptime(date, "%y%m%d")
    begin_date_obj = datetime.strptime(begin_date, "%y%m%d")
    end_date_obj = datetime.strptime(end_date, "%y%m%d")
    return begin_date_obj <= date_obj <= end_date_obj


# 호출: utils - get_sales_log()
# 동작: `재고 조사` 의 각 상품 데이터에 대해, `입고 Log`에서 가격 정보를 가져옴
# 주의: 대부분의 엑셀 시트 오류는 이 함수에서 catch
def get_sales_price(receive_log, product_name, date):
    for index, row in receive_log.iloc[::-1].iterrows():
        receive_date = datetime.strptime(str(row['일자']), '%y%m%d')
        # 띄어쓰기 무시하기 위함
        origin_product_name = re.sub(r"\s", "", row['상품명'])
        target_product_name = re.sub(r"\s", "", product_name)
        # 찾는 일자에서 가장 최신 일자의 상품 입고 데이터를 추출
        if origin_product_name == target_product_name and receive_date <= datetime.strptime(date, '%y%m%d'):
            sales_price = row['판매 금액']
            return '0' if pd.isna(sales_price) else sales_price
    # `재고 조사` 시트에는 존재하지만, `입고 Log` 시트에 존재하지 않는 경우에 대한 에러 처리
    print(f"{product_name}({date})는 존재하지 않는 상품입니다. 입고 Log를 다시 확인해 주십시오.")
    return


# 호출: process - straight_log 동작 함수들
# 동작: 해당 시트에서 출고된 모든 상품에 대해 가격 정보를 가져와 더해서 반환
def get_sales_log(stock_survey_sheet, date):
    sheet_list = get_sheet_list()
    receive_log = get_receive_log(sheet_list)
    sales_log = []
    for index, row in stock_survey_sheet.iterrows():
        if row['출고'] > 0:
            product_name = row['상품명']
            # 출고 * 판매 금액
            sales_profit = int(row['출고']) * int(get_sales_price(receive_log, product_name, date))
            sales_log.append({
                '종류': row['종류'],
                '상품명': product_name,
                '출고 수량': row['출고'],
                '판매 수입': sales_profit,
                '일자': date
            })
    return pd.DataFrame(sales_log)


# 호출: process - straight_log 동작 함수들
# 동작: 셀 병합으로 인해 `종류`가 None인 항목에 대해 종류를 기입
def set_product_kind_valid(stock_check_sheet):
    previous_value = None
    for index, row in stock_check_sheet.iterrows():
        if pd.isnull(row['종류']):
            stock_check_sheet.at[index, '종류'] = previous_value
        else:
            previous_value = row['종류']
    return stock_check_sheet


# 호출: process - get_expenses_straight_log
# 동작: `재고 조사`와 `입고 Log`의 정보를 합쳐 expenses_straight_log의 row 데이터를 구성
def get_expenses_log(row):
    return pd.DataFrame([{
        '종류': row['종류'],
        '상품명': row['상품명'],
        '입고 수량': row['입고 수량 (단품)'],
        '구매 비용': round(row['구매 비용']),
        '일자': row['일자']
    }])


# 호출: process - 엑셀 작성 함수들
# 동작: `종류` heads들을 엑셀에 작성
def append_product_kind_heads(ws, row_index):
    product_kind_heads = product_kind_list[:]
    product_kind_heads.append('합계')
    for index, product_kind in enumerate(product_kind_heads, start=1):
        col_index = index * 2
        ws.merge_cells(start_row=row_index, start_column=col_index, end_row=row_index, end_column=col_index + 1)
        ws.cell(row_index, col_index, product_kind)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: `종류` heads에 맞는 column들을 엑셀에 작성
def append_heads_columns(ws, row_index, column1, column2):
    product_kind_heads = product_kind_list[:]
    product_kind_heads.append('합계')
    for index, product_kind in enumerate(product_kind_heads, start=1):
        col_index = index * 2
        ws.cell(row_index, col_index, column1)
        ws.cell(row_index, col_index + 1, column2)
    return


# 호출: utils - 다수
# 동작: 해당 row (엑셀) 에서 `수량`, `가격` 데이터들을 정산하여 반환
def get_count_and_price_sum(row):
    # 수량 / 가격
    amount_sum = sum(row[1::2])
    price_sum = sum(row[2::2])
    return [amount_sum, price_sum]


# 호출: process - 엑셀 작성 함수들
# 동작: 각 일자 별 매출 데이터를 종합하여 엑셀에 작성
def insert_sales_data(ws, sales_report):
    date = ''
    row = []
    for index, data_row in sales_report.iterrows():
        # 같은 일자가 아니면 해당 일자의 데이터는 끝이므로 엑셀에 작성
        if date != data_row.name[0]:
            # 맨 처음 일자
            if row:
                row.extend(get_count_and_price_sum(row))
                ws.append(row)
            date = data_row.name[0]
            row = [date]
        # 같은 일자인 경우 추가
        row.extend([data_row['출고 수량'], data_row['판매 수입']])
    # 엑셀에 작성
    row.extend(get_count_and_price_sum(row))
    ws.append(row)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: 각 일자 별 매입 데이터를 종합하여 엑셀에 작성
def insert_expenses_data(ws, expenses_report):
    date = ''
    row = []
    for index, data_row in expenses_report.iterrows():
        # 같은 일자가 아니면 해당 일자의 데이터는 끝이므로 엑셀에 작성
        if date != data_row.name[0]:
            # 맨 처음 일자
            if row:
                row.extend(get_count_and_price_sum(row))
                ws.append(row)
            date = data_row.name[0]
            row = [date]
        # 같은 일자인 경우 추가
        row.extend([data_row['입고 수량'], data_row['구매 비용']])
    # 엑셀에 작성
    row.extend(get_count_and_price_sum(row))
    ws.append(row)
    return


# 호출: process - report 연산 함수들
# 동작: 모든 일자에 대해, 비어있는 `종류`에 0 데이터를 삽입
# 의도: 엑셀 기록시 `종류` 컨벤션을 맞춰 데이터가 밀리는 일이 없게 하기 위함
def append_kind_polyfill(sales_report, date_list):
    for date in date_list:
        for kind in product_kind_list:
            # 어차피 groupby sum 연산 수행하므로 0 데이터 삽입
            polyfill = pd.DataFrame(
                {'출고 수량': [0], '판매 수입': [0]},
                index=pd.MultiIndex.from_tuples(
                    [(date, kind)],
                    names=['일자', '종류']
                ))
            sales_report = pd.concat([sales_report, polyfill])
    return sales_report


# 호출: utils - insert_sum_data 함수들
# 동작: 비어있는 `종류`에 0 데이터를 삽입
# 의도: 엑셀 기록시 `종류` 컨벤션을 맞춰 데이터가 밀리는 일이 없게 하기 위함
def append_kind_polyfill_except_date(sales_sum_data):
    # 삽입을 위해 인덱스를 number로 초기화
    sales_sum_data.reset_index(inplace=True)
    for kind in product_kind_list:
        polyfill = pd.DataFrame({'종류': [kind], '출고 수량': [0], '판매 수입': [0]})
        sales_sum_data = pd.concat([sales_sum_data, polyfill], axis=0)
    return sales_sum_data


# 호출: process - 엑셀 작성 함수들
# 동작: 일자, 종류별 합계 데이터를 연산 후 엑셀에 작성
def insert_sales_sum_data(ws, sales_straight_log):
    sales_sum_data = sales_straight_log\
        .groupby(['종류'])\
        .agg({'출고 수량': 'sum', '판매 수입': 'sum'})\
        .sort_values(by='종류', ascending=True)
    # `종류` 컨벤션 맞춘 후 한번 더 연산
    sales_sum_data = append_kind_polyfill_except_date(sales_sum_data)
    sales_sum_data = sales_sum_data\
        .groupby(['종류'])\
        .agg({'출고 수량': 'sum', '판매 수입': 'sum'})\
        .sort_values(by='종류', ascending=True)
    row = ['합계']
    for index, sum_data in sales_sum_data.iterrows():
        row.extend([sum_data['출고 수량'], sum_data['판매 수입']])
    row.extend(get_count_and_price_sum(row))
    ws.append(row)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: 일자, 종류별 합계 데이터를 연산 후 엑셀에 작성
def insert_expenses_sum_data(ws, expenses_straight_log):
    expenses_sum_data = expenses_straight_log\
        .groupby(['종류'])\
        .agg({'입고 수량': 'sum', '구매 비용': 'sum'})\
        .sort_values(by='종류', ascending=True)
    # `종류` 컨벤션 맞춘 후 한번 더 연산
    expenses_sum_data = append_kind_polyfill_except_date(expenses_sum_data)
    expenses_sum_data = expenses_sum_data\
        .groupby(['종류'])\
        .agg({'입고 수량': 'sum', '구매 비용': 'sum'})\
        .sort_values(by='종류', ascending=True)
    row = ['합계']
    for index, sum_data in expenses_sum_data.iterrows():
        row.extend([sum_data['입고 수량'], sum_data['구매 비용']])
    row.extend(get_count_and_price_sum(row))
    ws.append(row)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: `순이익` 항목의 heads를 엑셀에 작성
def append_product_kind_short_heads(ws):
    row = [''] + product_kind_list + ['총합']
    ws.append(row)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: 상품 종류에 대해 순이익을 연산하여 엑셀에 작성
def insert_net_profit_row(ws, sales_row_count, expenses_row_count):
    row = ['순이익']
    # magic number 4, 9: 데이터를 제외한 heads row, 공백 row에 대한 개수 (고정값)
    sales_sum_row_index = 4 + sales_row_count
    expenses_sum_row_index = 9 + sales_row_count + expenses_row_count
    # 매출, 매입 합계 데이터가 기록된 row를 추출
    sales_sum_row = [cell.value for cell in ws[sales_sum_row_index]][2::2]
    expenses_sum_row = [cell.value for cell in ws[expenses_sum_row_index]][2::2]
    # 상품 종류에 대해, 종합에 대해 순이익 연산
    net_profit_row = [int(x - y) for x, y in zip(sales_sum_row, expenses_sum_row)]
    row.extend(net_profit_row)
    ws.append(row)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: 상품 랭킹에 필요한 heads를 엑셀에 작성
def append_sales_ranking_heads(ws):
    row = ['', '종류', '상품명', '출고 수량', '판매 수입']
    ws.append(row)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: 상품 랭킹을 단순 append하여 엑셀에 작성
def insert_sales_ranking_data(ws, product_sales_ranking):
    # 랭킹 삽입을 위해 인덱스를 number로 초기화
    product_sales_ranking.reset_index(inplace=True)
    # index가 랭킹
    for index, data_row in product_sales_ranking.iterrows():
        row = [index+1]
        row.extend(data_row.tolist())
        ws.append(row)
    return


# 호출: process - 엑셀 작성 함수들
# 동작: `종류`별 상품에 대해 매출 랭킹을 작성
def insert_sales_ranking_data_at_product(ws, product_sales_report):
    # 랭킹 삽입을 위해 인덱스를 number로 초기화
    product_sales_report.reset_index(inplace=True)
    ranking_index = 1
    kind = ''
    for index, data in product_sales_report.iterrows():
        # 다른 종류를 만나면 해당 종류는 끝이므로 엑셀 상으로 개행 후 랭킹 초기화
        if kind != data['종류']:
            ws.append([])
            kind = data['종류']
            ranking_index = 1
            ws.append(['', kind])
            ws.append([''] + product_sales_report.columns[1:].tolist())
        # 같은 종류라면 랭킹 append
        ws.append([ranking_index, data['상품명'], data['출고 수량'], data['판매 수입']])
        ranking_index += 1
    return


# 호출: process - 엑셀 작성 함수들
# 동작: `종류`별 상품에 대해 매입 랭킹을 작성
def insert_expenses_ranking_data_at_product(ws, product_expense_report):
    # 랭킹 삽입을 위해 인덱스를 number로 초기화
    product_expense_report.reset_index(inplace=True)
    ranking_index = 1
    kind = ''
    for index, data in product_expense_report.iterrows():
        # 다른 종류를 만나면 해당 종류는 끝이므로 엑셀 상으로 개행 후 랭킹 초기화
        if kind != data['종류']:
            ws.append([])
            kind = data['종류']
            ranking_index = 1
            ws.append(['', kind])
            ws.append([''] + product_expense_report.columns[2:].tolist())
        # 같은 종류라면 랭킹 append
        ws.append([ranking_index, data['상품명'], data['입고 수량'], data['구매 비용']])
        ranking_index += 1
    return
