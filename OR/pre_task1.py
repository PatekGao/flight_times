import pandas as pd

# 读取Excel文件
df = pd.read_excel('2024_5min_pax.xlsx', dtype={'Time': str})

ref_arr_dom = df['DOM_ARR'].values
ref_arr_int = df['INT_ARR'].values
ref_dep_dom = df['DOM_DEP'].values
ref_dep_int = df['INT_DEP'].values

# test
# print(ref_arr_dom)
# print(ref_arr_int)
# print(ref_dep_dom)
# print(ref_dep_int)
print(len(ref_arr_dom))
