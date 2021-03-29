#reference : https://towardsdatascience.com/learn-how-to-easily-do-3-advanced-excel-tasks-in-python-925a6b7dd081
import pandas as pd

if __name__ == "__main__":
    sales = pd.read_excel('https://github.com/datagy/mediumdata/raw/master/pythonexcel.xlsx', sheet_name = 'sales')
    states = pd.read_excel('https://github.com/datagy/mediumdata/raw/master/pythonexcel.xlsx', sheet_name = 'states')
    print(sales.head())
    sales['MoreThan500'] = ['Yes' if x > 500 else 'No' for x in sales['Sales']]
    # print(sales['MoreThan500'])

#VLOOKUP = 두개의 excel sheet를 같은 column 값을 key로 매핑해 left merge
    sales = pd.merge(sales, states, how='left', on='City')
    print(sales.head())

#pivot
    pivoted = sales.pivot_table(index = 'City', values = 'Sales', aggfunc = 'sum')
    print(pivoted.head())

"""
1. 각 월정액 상품별 총 시청시간 or 시청건수 : PPM > "월정액구분"을 index로 "시청시간(분)"이나 "시청건수"를 values로 pivot_table적용
2. 각 월정액 내 CP별 총 시청기여율 : PPM > "월정액구분" 리스트 중 1개씩 필터링 > 해당 월정액 내 "거래처명" 리스트 중 1개씩 필터링 > 시청시간이나 시청건수를 sum
3. 월정액 별 CP의 시청기여율 : 2에서 추출된 월정액 내 CP별 총 시청시간 / 1에서 추출된 각 월정액별 총 시청시간 x 100
여기에 1, 2, 3의 내용을 각 엑셀 back data로 별도 저장해 분류 제공 필요
"""