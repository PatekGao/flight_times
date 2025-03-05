import pandas as pd

# 读取Excel文件
df = pd.read_excel('TFU_dynamic_sheet2040_rand_smooth_2-23_ZG.xlsx', dtype={'Time': str})

records = []

for _, row in df.iterrows():
    time = row['Time']

    # 处理ARR组（到达航班）
    for _ in range(row['ARR_DOM']):
        records.append({
            'ID': '',
            '时间': time,
            '市场': 'DOM',
            '进出港': 'ARR',
            '机型': '',
            '日期': 2
        })
    for _ in range(row['ARR_INT']):
        records.append({
            'ID': '',
            '时间': time,
            '市场': 'INT',
            '进出港': 'ARR',
            '机型': '',
            '日期': 2
        })

    # 处理DEP组（出发航班）
    for _ in range(row['DEP_DOM']):
        records.append({
            'ID': '',
            '时间': time,
            '市场': 'DOM',
            '进出港': 'DEP',
            '机型': '',
            '日期': 2
        })
    for _ in range(row['DEP_INT']):
        records.append({
            'ID': '',
            '时间': time,
            '市场': 'INT',
            '进出港': 'DEP',
            '机型': '',
            '日期': 2
        })

# 创建完整DataFrame
full_df = pd.DataFrame(records)[['ID', '时间', '市场', '进出港', '机型', '日期']]

# 分组处理
arr_df = full_df[full_df['进出港'] == 'ARR']
dep_df = full_df[full_df['进出港'] == 'DEP']

# 统计日期为2的0-10点出港航班数量
dep_flights_2_0_10 = dep_df[(dep_df['日期'] == 2) & (dep_df['时间'].between('00:00', '10:00'))].shape[0]

# 国际出港航班数量
international_flights = \
    dep_df[(dep_df['日期'] == 2) & (dep_df['时间'].between('00:00', '10:00')) & (dep_df['市场'] == 'INT')].shape[0]

# 国内出港航班数量
domestic_flights = \
    dep_df[(dep_df['日期'] == 2) & (dep_df['时间'].between('00:00', '10:00')) & (dep_df['市场'] == 'DOM')].shape[0]

print(f"日期为2的0-10点出港航班数量: {dep_flights_2_0_10}")
print(f"其中国际出港航班数量: {international_flights}")
print(f"其中国内出港航班数量: {domestic_flights}")

day1_int_flights = round(international_flights * 0.6)
day1_dom_flights = round(domestic_flights * 0.6)
print(f'第一天国际航班量：{day1_int_flights}')
print(f'第一天国内航班量：{day1_dom_flights}')

late_int_flights = arr_df[arr_df['市场'] == 'INT'].sort_values('时间', ascending=False).head(day1_int_flights)
late_dom_flights = arr_df[arr_df['市场'] == 'DOM'].sort_values('时间', ascending=False).head(day1_dom_flights)
late_flights = pd.concat([late_int_flights, late_dom_flights])
late_flights['日期'] = 1

arr_df = pd.concat([late_flights, arr_df]).reset_index(drop=True)

# 保存到不同sheet
with pd.ExcelWriter('pre_task2.xlsx') as writer:
    arr_df.to_excel(writer, sheet_name='到达航班', index=False)
    dep_df.to_excel(writer, sheet_name='出发航班', index=False)
