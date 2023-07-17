# TF 교회 카페팀 엑셀 업무 자동화 프로젝트

</br>

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/b9ce6fb2-d276-428a-8a35-d759dc41af0b)

</br>

---

</br>

# 사용 방법

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/3b9e9dfa-1c82-46ba-b53f-a88ae6ab0902)

`main.py` 파일에서 초록색 실행 버튼을 클릭

</br>

---

</br>

# 주의 사항

</br>

🛑 입력 엑셀은 MicroSoft Excel 프로그램으로 작성되어야 작동함

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/3f91beef-5a3c-4804-8f28-24f0dc0c1e11)

</br>

🛑 모든 시트의 column 이름, 순서는 **절대로** 수정 금지

</br>

🛑 상품 수량, 단가, 판매 가격 등 모든 숫자 데이터는 0을 포함한 **양의 정수**로 작성
- 소수점 X, 음수 X, 빈칸 X

</br>

🛑 모든 날짜는 반드시 `yymmdd` 형태로 기록
- ex) 230630, 230715

</br>

🛑 `재고 조사` 시트의 상품 명과 `입고 Log` 시트의 상품 명이 다르면 오류 발생

</br>

🛑 `구매 리스트 템플릿 yymmdd`, `yymmdd w 재고 조사` **시트 이름** 양식을 지킬것
- ex) 구매 리스트 템플릿 230701
- ex) 230718 화 재고 조사

</br>

🛑 상품의 `종류`를 추가하고자 한다면 개발자에게 연락하여 코드 수정을 요청할것
- 현재 상품 `종류`는 아래와 같음
```
간식 / 라면, 인스턴트 / 생필품 / 아이스크림 / 음료 / 일회용품 /친환경제품 / 타먹는 차
```

</br>

---

</br>

# 입력 엑셀

</br>

## 🔸 1. 구매 템플릿

</br>

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/4bec8024-2358-416d-8158-9d0fd03a0a7b)

코드와 상관 없는 데이터로, 시트 이름 양식만 준수할것

</br>

## 🔸 2. 입고 Log

</br>

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/3baab03f-1815-4072-857b-72f9ec6ddc27)

### 기록 방법

```
같은 상품을 추가로 입고하는 경우, 기존 기록을 수정하지 말고 단순히 아래에 추가할것
상품 가격의 변동을 고려한 것이 의도이고,
일자가 다른 것을 인식하여 코드가 동작하기 때문에 괜찮음
```

</br>

`종류`, `상품명` 오타 주의

`입고 수량 (단품)`: 세트, 묶음이 아닌 단위 상품 수량

`단품 가격`: `구매 비용` / `입고 수량 (단품)`

`구매 비용`: 해당 상품을 총 구매하는 데 든 금액

`판매 금액`: 해당 상품을 판매하기 위해 책정한 금액

`일자`: 해당 상품을 입고한 일자

</br>

## 🔸 3. 재고 조사

</br>

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/cd991387-5857-4b16-99cf-58f2ee8e369e)

`이전 재고`: 바로 직전의 재고 조사 시트의 `현재 재고`

`현재 재고`: 기록 당시의 재고

`출고`, `입고`: 물품 수량을 **음이 아닌 정수**로 기록

`입고 일자`: 해당 상품이 입고된 일자

</br>

```
이미 지난 재고 조사에 입고 일자를 기록했다면 또 기록할 필요 없음
ex)
230710 생수 50개 입고
-> 230711 재고 조사: 생수 / 입고 일자 (230710)
-> 230713 재고 조사: 생수 / 입고 일자는 전에 이미 기록했으니 또 기록할 필요 없음
```

</br>

---

</br>

# 기능

</br>

## 1. 특정 구간 내 매출, 매입 데이터 산출

</br>

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/e052ebc2-bc61-4e54-b880-fd7ef4b7b57f)

데이터 산출을 원하는 시작 날짜, 종료 날짜를 기록하면 아래와 같이 시트가 2개 생성됨

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/038c6b65-2524-4d60-b33d-963e99396bd6)

</br>

### 🔸 overall report

**종류별 매출 종합**

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/9e6ad7a1-90a2-482b-a793-084a55646a7a)


**종류별 지출 종합**

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/4fd53bbb-680c-4589-bdcb-d7f357a1a837)


**종류별 순이익**

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/dbcbcd49-b92f-47ea-8804-9e44a6645bfb)


**상품 판매 순위**

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/d940f922-ff82-4105-83b0-e4d7e924fc10)

</br>

### 🔸 product report

**`종류`별 상품에 대한 매출 순위 **

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/5a952c69-de87-4ab8-a208-af918ac102ce)

</br>

## 2. 특정 요일의 매출, 매입 데이터 산출

</br>

미구현

</br>

---

</br>

# 호출 관계

</br>

![image](https://github.com/edac99/tf_cafe_excel/assets/79911816/cd5a2ae2-2005-477e-8ae7-8152c2baa69c)



