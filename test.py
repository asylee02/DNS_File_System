import pandas as pd
first = []
second = []
third = []
fourth = []
fifth = []
# 읽어올 엑셀 파일 지정
filename = 'DNS_Test.xlsx'

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

# 엑셀 파일 읽어 오기
df1 = pd.read_excel(filename, sheet_name=0, engine='openpyxl', usecols=[1,2,10] )
cleaned_df = df1.dropna(axis=0, how='any')
for index, row in cleaned_df.iterrows():
    row_values = row.values.tolist()
    first.append(row_values)


df2 = pd.read_excel(filename, sheet_name=1, engine='openpyxl', usecols=[2,6,8] )
cleaned_df = df2.dropna(axis=0, how='any')
for index, row in cleaned_df.iterrows():
    row_values = row.values.tolist()
    second.append(row_values)
second.pop(0)


df3 = pd.read_excel(filename, sheet_name=2, engine='openpyxl', usecols=[2,6] )
cleaned_df = df3.dropna(axis=0, how='any')
for index, row in cleaned_df.iterrows():
    row_values = row.values.tolist()
    third.append(row_values)
third.pop()
third.pop(0)


        
df4 = pd.read_excel(filename, sheet_name=4, engine='openpyxl', usecols=[1,5,7])
cleaned_df = df4.dropna(axis=0, how='any')
for index, row in cleaned_df.iterrows():
    row_values = row.values.tolist()
    fourth.append(row_values)
fourth.pop(0)
fourth = [remove_prefix(item) for item in fourth]

df5 = pd.read_excel(filename, sheet_name=5, engine='openpyxl', usecols=[2,5])
cleaned_df = df5.dropna(axis=0, how='any')
for index, row in cleaned_df.iterrows():
    row_values = row.values.tolist()
    fifth.append(row_values)
fifth = duplication(fifth)

for i in range(len(first)):
    answer = first[i]
    if first[i][1] != second[i][1]:
        print(f"PI 시트에서 {second[i][0]} 제품의 갯수가 틀렸습니다.")
    if first[i][2] != second[i][2]:
        print(f"PI 시트에서 {second[i][0]} 제품의 가격이 틀렸습니다.")
    if first[i][1] != third[i][1]:
        print(f"구매발주서(건흥) 시트에서 {third[i][0]} 제품의 갯수가 틀렸습니다.")
    if first[i][1] != fourth[i][1]:
        print(f"Invoice 시트에서 {third[i][0]} 제품의 갯수가 틀렸습니다.")
    if first[i][2] != fourth[i][2]:
        print(f"Invoice 시트에서 {fourth[i][0]} 제품의 가격이 틀렸습니다.")
    if first[i][1] != fifth[i][1]:
        print(f"Packing 시트에서 {fifth[i][0]} 제품의 갯수가 틀렸습니다.")
    

#     file_path = 'DNS_Test.xlsx' # 파일 경로를 적절히 변경하세요
# excel_data = pd.ExcelFile(file_path)

# # 시트의 개수와 이름 확인
# sheet_names = excel_data.sheet_names
# num_sheets = len(sheet_names)

# print(f"시트의 개수: {num_sheets}")
# print("각 시트의 이름:")
# for idx, sheet_name in enumerate(sheet_names):
#     print(f"시트 인덱스 {idx}: {sheet_name}")
