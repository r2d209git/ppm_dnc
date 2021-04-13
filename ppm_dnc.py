#reference : https://towardsdatascience.com/learn-how-to-easily-do-3-advanced-excel-tasks-in-python-925a6b7dd081
import pandas as pd

#개별 컨텐츠의 sheet 작성
def write_each_ppm(ppmAll, ppm):
    product = ppmAll[isMatched].sort_values(by=['시청시간(분)'], axis=0, ascending=False)
    totalHours1 = product['시청시간(분)'].sum()
    product['점유율'] = (product['시청시간(분)'] / totalHours1) * 100

    #필요 output column만 추출
    df_select = product.loc[:, ['NCMS카테고리명', '거래처명', '타이틀', '시청회차', '단가', '시청시간(분)', '점유율']]
    #column renaming
    df_select.columns = ['카테고리', 'CP명', 'TITLE', '회차', '단가', '시청시간(분)', '점유율']
    #column reordering
    df_select = df_select[['카테고리', 'TITLE', '회차', '시청시간(분)', '점유율', 'CP명', '단가']]

    df_select.to_excel(writer_each, sheet_name = ppm)

    return product

#월정액 총합 sheet 작성
def write_total_ppm(product, ppm):
    hoursPerCP = product.pivot_table(index = '거래처명', values = '시청시간(분)', aggfunc = 'sum')
    totalHours2 = hoursPerCP['시청시간(분)'].sum()
    hoursPerCP['점유율'] = (hoursPerCP['시청시간(분)'] / totalHours2) * 100
    hoursPerCP = hoursPerCP.sort_values(by=['점유율'], axis=0, ascending=False)

    #필요 output column만 추출
    # df_select = hoursPerCP.loc[:, ['거래처명', '점유율']]
    #형식대로 column name renaming
    # df_select.columns = ['CP명', '점유율']
    hoursPerCP.to_excel(writer_each, sheet_name = ppm+'_분류')

if __name__ == "__main__":
    ppmAll = pd.read_excel('./Btv_PPM_시청시간_202103.xlsx', sheet_name = '전체')
    # states = pd.read_excel('https://github.com/datagy/mediumdata/raw/master/pythonexcel.xlsx', sheet_name = 'states')
    # print(ppmAll.head())

# 1-1. 각 월정액 상품별 총 시청시간 : PPM > "월정액구분"을 index로 "시청시간(분)"를 values로 pivot_table적용
    hoursPerPPM = ppmAll.pivot_table(index = '월정액구분', values = '시청시간(분)', aggfunc = 'sum')
    # print(hoursPerPPM.index)
    # print(hoursPerPPM)
    

# 1-2 각 월정액 상품별 총 시청건수 : PPM > "월정액구분"을 index로 "시청건수"를 values로 pivot_table적용
    viewPerPPM = ppmAll.pivot_table(index = '월정액구분', values = '시청건수', aggfunc = 'sum')
    # print(viewPerPPM)

# 2. 각 월정액 내 CP별 총 시청기여율 : PPM > "월정액구분" 리스트 중 1개씩 필터링 > 해당 월정액 내 "거래처명" 리스트 중 1개씩 필터링 > 시청시간이나 시청건수를 sum
# 2-1. PPM > "월정액구분" 리스트 중 1개씩 필터링
    with pd.ExcelWriter('./Btv_PPM_시청시간_개별.xlsx') as writer_each:
        for ppm in hoursPerPPM.index:
            print(f'Product = {ppm}')
            isMatched = (ppmAll['월정액구분'] == ppm)
#2-2. 해당 월정액 내 "거래처명" 리스트 중 1개씩 필터링 > 시청시간(이나 시청건수를) sum
            if ppm == "OCEAN" or ppm == "JTBC+OCEAN" or ppm == "OCEAN M" or ppm == "19월정액" or ppm == "해피시니어 월정액":
                product = write_each_ppm(ppmAll, ppm)
                write_total_ppm(product, ppm)               


# 3. 월정액 별 CP의 시청기여율 : 2에서 추출된 월정액 내 CP별 총 시청시간 / 1에서 추출된 각 월정액별 총 시청시간 x 100
    with pd.ExcelWriter('./Btv_PPM_시청시간_종합.xlsx') as writer_total:
        hoursPerPPM.to_excel(writer_total, sheet_name = '월정액상품별총시청시간')  
        viewPerPPM.to_excel(writer_total, sheet_name = '월정액상품별총시청건수')   



    # sales['MoreThan500'] = ['Yes' if x > 500 else 'No' for x in sales['Sales']]
    # print(sales['MoreThan500'])

#VLOOKUP = 두개의 excel sheet를 같은 column 값을 key로 매핑해 left merge
#     sales = pd.merge(sales, states, how='left', on='City')
#     print(sales.head())