import pandas as pd
from openpyxl import load_workbook

file_path = 'Test.xlsx' # 파일 경로를 적절히 변경하세요

# wb = load_workbook(filename=file_path)

# # 숨겨진 시트 이름 확인
# hidden_sheets = [sheet.title for sheet in wb.worksheets if sheet.sheet_state == 'hidden']

# # 숨겨진 시트 삭제
# for sheet_name in hidden_sheets:
#     del wb[sheet_name]

# # 변경된 내용을 적용하고 파일로 저장
# wb.save('Test.xlsx')

excel_data = pd.ExcelFile(file_path)
pd.read_excel(file_path,  engine='openpyxl' )
sheet = []
# 시트의 개수와 이름 확인
sheet_names = excel_data.sheet_names
num_sheets = len(sheet_names) -1
print(f"시트의 개수: {num_sheets}")
for idx, sheet_name in enumerate(sheet_names):
    sheet.append(sheet_name)
    print(sheet_name)





total_cnt = 0

Total_name_List =[]

# 추가 PI
def plus_Proforma(i):
    globals()[f'List{i-1}'] = []
    df1 = pd.read_excel(file_path, sheet_name=i, engine='openpyxl', usecols=[1,2,10] )
    cleaned_df = df1.dropna(axis=0, how='any')
    for index, row in cleaned_df.iterrows():
        row_values = row.values.tolist()
        globals()[f'List{i-1}'].append(row_values)

# 추가 발주서
def plus_Order(i):
    globals()[f'List{i-1}'] = []
    df1 = pd.read_excel(file_path, sheet_name=i, engine='openpyxl', usecols=[2,6] )
    cleaned_df = df1.dropna(axis=0, how='any')
    for index, row in cleaned_df.iterrows():
        row_values = row.values.tolist()
        globals()[f'List{i-1}'].append(row_values)
    globals()[f'List{i-1}'].pop()
    globals()[f'List{i-1}'].pop(0)
    

def remove_prefix(item):
    if isinstance(item[0], str):
        return [item[0].split('. ', 1)[1]] + item[1:]
    else:
        return item

def duplication(file):
    result_dict={}
    for item in file:
        if item[0] in result_dict:
            result_dict[item[0]] += item[1]
        else:
            result_dict[item[0]] = item[1]
    new_list = [[key, value] for key, value in result_dict.items()]
    return new_list

def Extraction(i, List):
    globals()[f'List{i}'] = []
    df1 = pd.read_excel(file_path, sheet_name=i, engine='openpyxl', usecols=List )
    cleaned_df = df1.dropna(axis=0, how='any')
    for index, row in cleaned_df.iterrows():
        row_values = row.values.tolist()
        globals()[f'List{i}'].append(row_values)    

# 단가
def Price(i):
    Extraction(i,[1,2])
# PI
def Proforma(i):
    Extraction(i,[2,6,8])
    globals()[f'List{i}'].pop(0)

# 발주서
def Order(i):
    Extraction(i,[2,6])
    globals()[f'List{i}'].pop()
    globals()[f'List{i}'].pop(0)
 
# CI
def Commercial(i):
    Extraction(i,[1,5,7])
    delete_list = []
    for k in range(len(globals()[f'List{i}'])):
        if globals()[f'List{i}'][k][0] == 'Description of Goods':
            delete_list.append(k)
    delete_list.reverse()
    for j in delete_list:
        globals()[f'List{i}'].pop(j)
    globals()[f'List{i}'] = [remove_prefix(item) for item in globals()[f'List{i}']]

def Packing(i):
    Extraction(i,[2,5])
    globals()[f'List{i}'] = duplication(globals()[f'List{i}'])

def Compare(List_a, List_b,i):
    print(f'{sheet[i]} 탐색중')
    cnt = len(List_b)
    priceIs = len(List_b[0])
    for i in range(cnt):
        name = List_b[i][0]
        for j in range(len(List_a)):
            if List_a[j][0] == name:
                print(name)
                print(List_a)
                if priceIs==3:
                    if List_a[j][1] != List_b[i][1]:
                        print(f'{sheet[i]} 시트에서 {List_b[i][0]} 제품의 수량이 틀렸습니다')
                    if List_a[j][2] != List_b[i][2]:
                        print(f'{sheet[i]} 시트에서 {List_b[i][0]} 제품의 가격이 틀렸습니다')
                else:
                    if List_a[j][1] != List_b[i][1]:
                        print(f'{sheet[i]} 시트에서 {List_b[i][0]} 제품의 수량이 틀렸습니다')


for i in range(num_sheets):  
    total_cnt+=1
    # print(f'{sheet[i]} 검사')
    if '단가' in sheet[i]:
        Price(i)
    elif 'PI' in sheet[i]:
        if 'SHIPPING MARK' == sheet[i]:
            continue
        if '추가' in sheet[i]:
            total_cnt-=1
            plus_Proforma(i)
        else:
            Proforma(i)
    elif '발주서' in sheet[i]:
        if '추가' in sheet[i]:
            total_cnt-=1
            plus_Proforma(i)
        else:
            Order(i)
    elif 'CI' in sheet[i].upper() or 'INVOICE' in sheet[i].upper():
        Commercial(i)
    elif 'PL' in sheet[i].upper() or 'PACKING' in sheet[i].upper():
        Packing(i)
    else:
        print('-----------------------------------------')
        print('--------시트 명을 다시 확인해주세요--------')
        print('-----------------------------------------')

for j in range(len(List0)):
    print(j)
    print(List0)
    print(List1)
    Total_name_List.append(List0[j][0])
    if List0[j][1] != List1[j][1]:
        print(f'PI시트에서 {List1[j][0]} 제품의 수량이 틀렸음')
    if List0[j][2] != List1[j][2]:
        print(f'PI시트에서 {List1[j][0]} 제품의 가격이 틀렸음')


for i in range(num_sheets-2):
    Compare(List0, globals()[f'List{i+2}'],i+2 )
