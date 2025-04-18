import pandas as pd

# 读取Excel文件
df = pd.read_excel('2024_现状.xlsx')

df = df[df['Date'] == 2]
df = df[df['flight type'] == 'PAX']

# 筛选到达航班
df = df[df['Direction'] == 'Arrival']
# df = df[df['Direction'] == 'Departure']

# 获取所有可能的航司、航向和机型
all_airlines = sorted(df['AirlineGroup'].unique())
all_routings = sorted(df['Routing'].unique())
all_acft_cats = sorted(df['Acft Cat'].unique())

# 创建所有可能的组合
all_combinations = pd.MultiIndex.from_product(
    [all_airlines, all_routings, all_acft_cats],
    names=['AirlineGroup', 'Routing', 'Acft Cat']
)

# 创建所有24小时
all_hours = list(range(24))

# 将时间列转换为小时格式
df['hour'] = df['time'].apply(lambda x: x.hour)

# 创建Excel写入器
with pd.ExcelWriter('到达航班比例表.xlsx') as writer:
# with pd.ExcelWriter('出发航班比例表.xlsx') as writer:
    # 处理国内航班
    df_domestic = df[df['flightnature'] == 'DOM']

    # 创建国内航班透视表
    pivot_domestic = pd.pivot_table(
        df_domestic,
        values='ID',
        index=['AirlineGroup', 'Routing', 'Acft Cat'],
        columns='hour',
        aggfunc='count',
        fill_value=0
    )

    # 重新索引国内航班透视表
    pivot_domestic = pivot_domestic.reindex(all_combinations, fill_value=0)
    pivot_domestic = pivot_domestic.reindex(columns=all_hours, fill_value=0)

    # 计算国内航班每个小时的总数
    pivot_domestic = pivot_domestic.div(509) * 100

    # 处理国际航班
    df_international = df[df['flightnature'] == 'INT']

    # 创建国际航班透视表
    pivot_international = pd.pivot_table(
        df_international,
        values='ID',
        index=['AirlineGroup', 'Routing', 'Acft Cat'],
        columns='hour',
        aggfunc='count',
        fill_value=0
    )

    # 重新索引国际航班透视表
    pivot_international = pivot_international.reindex(all_combinations, fill_value=0)
    pivot_international = pivot_international.reindex(columns=all_hours, fill_value=0)

    # 计算国际航班每个小时的总数
    pivot_international = pivot_international.div(52) * 100

    # 对列进行排序（按小时）
    pivot_domestic = pivot_domestic.reindex(sorted(pivot_domestic.columns), axis=1)
    pivot_international = pivot_international.reindex(sorted(pivot_international.columns), axis=1)

    # 保存到不同的sheet
    pivot_domestic.to_excel(writer, sheet_name='国内航班')
    pivot_international.to_excel(writer, sheet_name='国际航班')

print("出发航班比例表已生成完成，请查看'到达航班比例表.xlsx'文件。")
# print("出发航班比例表已生成完成，请查看'出发航班比例表.xlsx'文件。")
