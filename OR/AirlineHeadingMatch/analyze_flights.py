import pandas as pd

def process_flight_data(df, direction):
    # 筛选特定方向的航班并创建副本
    df_direction = df[df['Direction'] == direction].copy()
    print(f"\n处理{direction}航班数据:")
    print(f"筛选后的航班总数: {len(df_direction)}")
    print("\n数据示例:")
    print(df_direction[['AirlineGroup', 'Routing', 'Acft Cat', 'Minute for rolling', 'flightnature']].head())
    
    # 获取所有可能的航司、航向和机型
    all_airlines = sorted(df_direction['AirlineGroup'].unique())
    all_routings = sorted(df_direction['Routing'].unique())
    all_acft_cats = sorted(df_direction['Acft Cat'].unique())
    
    print(f"\n航司列表: {all_airlines}")
    print(f"航向列表: {all_routings}")
    print(f"机型列表: {all_acft_cats}")
    
    # 创建所有可能的组合
    all_combinations = pd.MultiIndex.from_product(
        [all_airlines, all_routings, all_acft_cats],
        names=['AirlineGroup', 'Routing', 'Acft Cat']
    )
    
    # 创建所有24小时
    all_hours = list(range(24))
    
    # 将分钟数转换为小时格式（确保在0-23范围内）
    df_direction.loc[:, 'hour'] = (df_direction['Minute for rolling'] // 60) % 24
    print("\n小时分布:")
    print(df_direction['hour'].value_counts().sort_index())
    
    # 处理国内航班
    df_domestic = df_direction[df_direction['flightnature'] == 'DOM'].copy()
    # 计算国内航班总数
    total_domestic = len(df_domestic)
    print(f"\n国内航班总数: {total_domestic}")
    if total_domestic > 0:
        print("\n国内航班示例:")
        print(df_domestic[['AirlineGroup', 'Routing', 'Acft Cat', 'hour']].head())
    
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
    if total_domestic > 0:
        pivot_domestic = pivot_domestic.div(total_domestic) * 100
        print("\n国内航班透视表示例:")
        print(pivot_domestic.head())
    else:
        print("警告：国内航班总数为0，无法计算百分比")
    
    # 处理国际航班
    df_international = df_direction[df_direction['flightnature'] == 'INT'].copy()
    # 计算国际航班总数
    total_international = len(df_international)
    print(f"\n国际航班总数: {total_international}")
    if total_international > 0:
        print("\n国际航班示例:")
        print(df_international[['AirlineGroup', 'Routing', 'Acft Cat', 'hour']].head())
    
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
    if total_international > 0:
        pivot_international = pivot_international.div(total_international) * 100
        print("\n国际航班透视表示例:")
        print(pivot_international.head())
    else:
        print("警告：国际航班总数为0，无法计算百分比")
    
    # 对列进行排序（按小时）
    pivot_domestic = pivot_domestic.reindex(sorted(pivot_domestic.columns), axis=1)
    pivot_international = pivot_international.reindex(sorted(pivot_international.columns), axis=1)
    
    return pivot_domestic, pivot_international

def main():
    # 读取Excel文件
    print("开始读取Excel文件...")
    input_name = '2024_现状_WUH.xlsx'
    df = pd.read_excel(input_name)
    print(f"原始数据总行数: {len(df)}")
    print("\n原始数据示例:")
    print(df[['Direction', 'flight type', 'Date', 'flightnature', 'Minute for rolling']].head())
    
    # 筛选数据
    df = df[df['Date'] == 2].copy()
    print(f"\n筛选Date=2后的行数: {len(df)}")
    
    df = df[df['flight type'] == 'PAX'].copy()
    print(f"筛选flight type=PAX后的行数: {len(df)}")
    
    # 处理出发航班
    pivot_domestic_dep, pivot_international_dep = process_flight_data(df, 'Departure')
    with pd.ExcelWriter(f'出发航班比例表{input_name}.xlsx') as writer:
        pivot_domestic_dep.reset_index().to_excel(writer, sheet_name='国内航班', index=False)
        pivot_international_dep.reset_index().to_excel(writer, sheet_name='国际航班', index=False)
    
    # 处理到达航班
    pivot_domestic_arr, pivot_international_arr = process_flight_data(df, 'Arrival')
    with pd.ExcelWriter(f'到达航班比例表{input_name}.xlsx') as writer:
        pivot_domestic_arr.reset_index().to_excel(writer, sheet_name='国内航班', index=False)
        pivot_international_arr.reset_index().to_excel(writer, sheet_name='国际航班', index=False)
    
    print(f"\n出发航班比例表{input_name}和到达航班比例表{input_name}已生成完成，请查看相应的Excel文件。")

if __name__ == "__main__":
    main()
