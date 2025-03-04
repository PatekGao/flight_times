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

# 保存到不同sheet
with pd.ExcelWriter('pre_task2.xlsx') as writer:
    arr_df.to_excel(writer, sheet_name='到达航班', index=False)
    dep_df.to_excel(writer, sheet_name='出发航班', index=False)
