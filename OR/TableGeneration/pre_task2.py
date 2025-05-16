from datetime import timedelta, datetime

import pandas as pd

xls = pd.ExcelFile("python时刻表输入v3.xlsx")
df2 = pd.read_excel(xls, sheet_name="其它参数设置")
DOM_MIX_p = df2.iloc[20, 3]
INT_MIX_p = df2.iloc[20, 4]

# 新增功能：统计第二天各类别高峰时段
def find_peak_window(df, market, direction):
    """找出指定市场和方向的高峰时段"""
    df = df[(df['市场'] == market) & (df['日期'] == 2)]
    time_counts = []

    # 生成所有有效时间窗口
    for start_time in pd.date_range("00:00", "23:59", freq="5T").time:
        start_min = start_time.hour * 60 + start_time.minute
        if start_min + 55 > 1439:  # 过滤跨日窗口
            continue

        # 计算时间窗口
        start_str = start_time.strftime("%H:%M")
        end_time = (datetime.combine(datetime.today(), start_time)
                    + timedelta(minutes=55)).time().strftime("%H:%M")
        window = f"{start_str}-{end_time}"

        # 统计该窗口内的航班数
        mask = df['时间'].between(start_str, end_time)
        time_counts.append((window, mask.sum()))

    # 找出最高峰窗口
    return max(time_counts, key=lambda x: x[1]) if time_counts else ("N/A", 0)


# 为每个类别计算高峰时段（使用向量化操作优化性能）
def add_minutes_column(df):
    """为DataFrame添加分钟数列"""
    df = df.copy()
    split_time = df['时间'].str.split(":", expand=True).astype(int)
    df['minutes'] = split_time[0] * 60 + split_time[1]
    return df


# 定义统计函数
def count_in_window(start_min, market, direction):
    """统计指定时间窗口的航班数"""
    end_min = start_min + 55
    if direction == "ARR":
        df = arr_df_day2[arr_df_day2['市场'] == market]
    elif direction == "DEP":
        df = dep_df_day2[dep_df_day2['市场'] == market]
    return ((df['minutes'] >= start_min) & (df['minutes'] <= end_min)).sum()


# 读取Excel文件
# df = pd.read_excel('TFU_dynamic_sheet2040_rand_smooth_2-23_ZG.xlsx', dtype={'Time': str})
df = pd.read_excel('dynamic_sheet.xlsx', dtype={'Time': str})
# 各类型高峰时段统计：
# 国内双向    : 13:20-14:15 (119 架次)(0.84%)
# 国内进港    : 13:30-14:25 (71 架次)(1.41%)
# 国内离港    : 07:35-08:30 (77 架次)(1.3%)
# 国际双向    : 01:50-02:45 (26 架次)(3.85%)
# 国际进港    : 22:25-23:20 (16 架次)(6.25%)
# 国际离港    : 09:05-10:00 (18 架次)(5.56%)
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

day1_int_flights = round(international_flights * INT_MIX_p)
day1_dom_flights = round(domestic_flights * DOM_MIX_p)
print(f'第一天国际航班量：{day1_int_flights}')
print(f'第一天国内航班量：{day1_dom_flights}')

late_int_flights = arr_df[arr_df['市场'] == 'INT'].sort_values('时间', ascending=False).head(day1_int_flights)
late_dom_flights = arr_df[arr_df['市场'] == 'DOM'].sort_values('时间', ascending=False).head(day1_dom_flights)
late_flights = pd.concat([late_int_flights, late_dom_flights])
late_flights['日期'] = 1

arr_df = pd.concat([late_flights, arr_df]).reset_index(drop=True)

# 预处理数据
arr_df_day2 = add_minutes_column(arr_df[arr_df['日期'] == 2])
dep_df_day2 = add_minutes_column(dep_df[dep_df['日期'] == 2])

# 主分析逻辑
categories = {
    "国内双向": ("DOM", "both"),
    "国内进港": ("DOM", "ARR"),
    "国内离港": ("DOM", "DEP"),
    "国际双向": ("INT", "both"),
    "国际进港": ("INT", "ARR"),
    "国际离港": ("INT", "DEP")
}

peak_results = {}
for cat_name, (market, direction) in categories.items():
    max_count = 0
    best_window = ""

    # 遍历所有时间窗口
    for start_time in pd.date_range("00:00", "23:59", freq="5min").time:
        start_min = start_time.hour * 60 + start_time.minute
        if start_min + 55 > 1439:
            continue

        # 计算窗口航班数
        if direction == "both":
            count = count_in_window(start_min, market, "ARR") + count_in_window(start_min, market, "DEP")
        else:
            count = count_in_window(start_min, market, direction)

        # 更新最大值
        if count > max_count or (count == max_count and best_window == ""):
            max_count = count
            window_start = start_time.strftime("%H:%M")
            window_end = (datetime.combine(datetime.today(), start_time)
                          + timedelta(minutes=55)).time().strftime("%H:%M")
            best_window = f"{window_start}-{window_end}"

    peak_results[cat_name] = (best_window, max_count)

# 打印结果
print("\n各类型高峰时段统计：")
for category, (window, count) in peak_results.items():
    print(f"{category.ljust(8)}: {window} ({count} 架次)({round((1 / count) * 100, 2)}%)")

# 保存到不同sheet
with pd.ExcelWriter('pre_task2.xlsx') as writer:
    arr_df.to_excel(writer, sheet_name='到达航班', index=False)
    dep_df.to_excel(writer, sheet_name='出发航班', index=False)
