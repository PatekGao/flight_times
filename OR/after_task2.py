from collections import defaultdict

import numpy as np
import pandas as pd
from gurobipy import Model, GRB, quicksum


# 预处理函数：将时间字符串转换为分钟数
def time_to_minutes(t_str):
    h, m = map(int, t_str.split(':'))
    return h * 60 + m  # 机型和市场配额约束（日期均为2）


quotas = {
    ('DOM', 'C'): {'ARR': 779, 'DEP': 779},
    ('DOM', 'E'): {'ARR': 87, 'DEP': 87},
    ('DOM', 'F'): {'ARR': 0, 'DEP': 0},
    ('INT', 'C'): {'ARR': 116, 'DEP': 115},
    ('INT', 'E'): {'ARR': 61, 'DEP': 61},
    ('INT', 'F'): {'ARR': 1, 'DEP': 1}
}

# 最短过站时间（分钟）
min_times = {
    ('DOM', 'C'): 35,
    ('DOM', 'E'): 50,
    ('DOM', 'F'): 50,
    ('INT', 'C'): 40,
    ('INT', 'E'): 55,
    ('INT', 'F'): 55
}

# 定义小时分布限制数据 (hour DOM_DEP INT_DEP)
dep_hour_distribution = [
    (0, 0, 1), (1, 0, 8), (2, 0, 2), (3, 0, 2), (4, 0, 1), (5, 0, 1),
    (6, 3, 0), (7, 6, 1), (8, 8, 2), (9, 4, 6), (10, 4, 2), (11, 6, 0),
    (12, 6, 1), (13, 6, 5), (14, 2, 3), (15, 6, 3), (16, 5, 6), (17, 8, 4),
    (18, 3, 4), (19, 5, 1), (20, 8, 2), (21, 3, 1), (22, 3, 1), (23, 1, 5)
]

peak_configs = [
    # 国内双向
    {
        'name': '国内双向',
        'market': 'DOM',
        'direction': 'both',
        'start_time': '13:20',
        'end_time': '14:15',
        'total': 119,
        'ratios': {'C': 0.88, 'E': 0.12, 'F': 0}
    },
    # 国内进港
    {
        'name': '国内进港',
        'market': 'DOM',
        'direction': 'ARR',
        'start_time': '13:30',
        'end_time': '14:25',
        'total': 71,
        'ratios': {'C': 0.88, 'E': 0.12, 'F': 0}
    },
    # 国内离港
    {
        'name': '国内离港',
        'market': 'DOM',
        'direction': 'DEP',
        'start_time': '07:35',
        'end_time': '08:30',
        'total': 77,
        'ratios': {'C': 0.88, 'E': 0.12, 'F': 0}
    },
    # 国际双向
    {
        'name': '国际双向',
        'market': 'INT',
        'direction': 'both',
        'start_time': '01:50',
        'end_time': '02:45',
        'total': 26,
        'ratios': {'C': 0.40, 'E': 0.60, 'F': 0}
    },
    # 国际进港
    {
        'name': '国际进港',
        'market': 'INT',
        'direction': 'ARR',
        'start_time': '22:25',
        'end_time': '23:20',
        'total': 16,
        'ratios': {'C': 0.395, 'E': 0.555, 'F': 0.05}
    },
    # 国际离港
    {
        'name': '国际离港',
        'market': 'INT',
        'direction': 'DEP',
        'start_time': '09:05',
        'end_time': '10:00',
        'total': 18,
        'ratios': {'C': 0.395, 'E': 0.555, 'F': 0.05}
    }
]

# 读取Excel文件中的到达和出发航班数据
arr_df = pd.read_excel('output_flight_pairing.xlsx', sheet_name='到达航班')
dep_df = pd.read_excel('output_flight_pairing.xlsx', sheet_name='出发航班')

# 筛选待配对的到达航班（ID为空且日期为2）
unpaired_arr = arr_df[(arr_df['ID'].isna()) & (arr_df['日期'] == 2)]

# 生成日期3的离港航班（复制日期2的离港航班）
date3_dep_df = dep_df[dep_df['日期'] == 2].copy()
date3_dep_df['日期'] = 3
date3_dep_df['ID'] = np.nan
date3_dep_df['机型'] = np.nan

# 合并离港航班数据
combined_dep_df = pd.concat([dep_df, date3_dep_df], ignore_index=True)


# 解析航班数据
def parse_flights(df, direction):
    return [
        {
            'index': idx,
            'time': time_to_minutes(row['时间']),
            'market': row['市场'],
            'date': row['日期'],
            'direction': direction
        }
        for idx, row in df.iterrows()
    ]


arr_flights = parse_flights(unpaired_arr, 'ARR')
dep_flights = parse_flights(combined_dep_df, 'DEP')

# 初始化模型
model = Model('Extended_Flight_Pairing')

# 生成变量：允许日期2到达与日期3出发配对
variables = {}
delta_dict = {}  # 新增：存储每个变量的过站时间
M = 2000  # 足够大的常数，需大于最大可能delta（24*60=1440）

for i, arr in enumerate(arr_flights):
    for j, dep in enumerate(dep_flights):
        if arr['date'] == 2 and dep['date'] == 3:
            delta = (24*60 - arr['time']) + dep['time']

            arr_market = arr['market']
            dep_market = dep['market']

            for k in ['C', 'E', 'F']:
                arr_quota = quotas.get((arr_market, k), {'ARR': 0})['ARR']
                min_time = max(
                    min_times.get((arr_market, k), 0),
                    min_times.get((dep_market, k), 0)
                )

                if delta >= min_time and arr_quota > 0:
                    var = model.addVar(vtype=GRB.BINARY, name=f'x_{i}_{j}_{k}')
                    variables[(i, j, k)] = var
                    delta_dict[(i, j, k)] = delta  # 记录过站时间

for i, arr in enumerate(arr_flights):
    for j, dep in enumerate(dep_flights):
        # 仅允许跨天配对（日期2到达 -> 日期3出发）
        if arr['date'] == 2 and dep['date'] == 3:
            # 计算过站时间（跨天）
            delta = (24 * 60 - arr['time']) + dep['time']

            # 获取市场信息
            arr_market = arr['market']
            dep_market = dep['market']

            # 检查机型约束（仅约束到达航班的配额）
            for k in ['C', 'E', 'F']:
                # 检查到达航班的配额是否允许
                arr_quota = quotas.get((arr_market, k), {'ARR': 0})['ARR']

                min_time = min_times.get((arr_market, k), 0)

                if delta >= min_time and arr_quota > 0:
                    variables[(i, j, k)] = model.addVar(vtype=GRB.BINARY, name=f'x_{i}_{j}_{k}')

# 目标函数：最大化 (M - delta) 的总和
valid_pairs = [(key, var) for key, var in variables.items() if key in delta_dict]
model.setObjective(
    quicksum( (M - delta_dict[key]) * var for key, var in valid_pairs ),
    GRB.MAXIMIZE
)
# 约束1：每个到达航班最多配对一次
arr_used = defaultdict(list)
for (i, j, k), var in variables.items():
    arr_used[i].append(var)
for i in arr_used:
    model.addConstr(quicksum(arr_used[i]) <= 1, name=f'ext_arr_{i}')

# 约束2：每个出发航班最多配对一次
dep_used = defaultdict(list)
for (i, j, k), var in variables.items():
    dep_used[j].append(var)
for j in dep_used:
    model.addConstr(quicksum(dep_used[j]) <= 1, name=f'ext_dep_{j}')

# 约束3：日期2的机型配额限制（仅约束到达航班）
for (market, k), quota in quotas.items():
    arr_total = [
        var for (i, j, key_k), var in variables.items()
        if key_k == k and arr_flights[i]['market'] == market
    ]
    tmp_arr = arr_df[(arr_df['市场'] == market) & (arr_df['机型'] == k) & (arr_df['日期'] == 2)]
    if arr_total:
        model.addConstr(quicksum(arr_total) + len(tmp_arr) <= quota['ARR'], f'ext_arr_quota_{market}_{k}')

# 约束4：日期2的高峰小时约束
for config in peak_configs:
    # 时间转换
    start_h, start_m = map(int, config['start_time'].split(':'))
    end_h, end_m = map(int, config['end_time'].split(':'))
    start_min = start_h * 60 + start_m
    end_min = end_h * 60 + end_m

    # 筛选到达航班
    selected = [
        i for i, arr in enumerate(arr_flights)
        if arr['market'] == config['market']
           and start_min <= arr['time'] <= end_min
    ]

    # 收集变量
    k_vars = defaultdict(list)
    for (i, j, k), var in variables.items():
        if i in selected:
            k_vars[k].append(var)

    # 添加约束
    ratios = config['ratios']
    total = config['total']
    for k in ['C', 'E', 'F']:
        if k in ratios and k_vars.get(k):
            required = round(total * ratios[k])
            model.addConstr(quicksum(k_vars[k]) <= required, f"ext_peak_{config['name']}_{k}")

# 求解模型
model.optimize()

# 处理配对结果
pair_id = max(arr_df['ID'].max(), dep_df['ID'].max()) + 1
pair_num = 0
for (i, j, k), var in variables.items():
    if var.X > 0.5:
        # 更新到达航班
        arr_idx = arr_flights[i]['index']
        arr_df.at[arr_idx, 'ID'] = pair_id
        arr_df.at[arr_idx, '机型'] = k

        # 更新出发航班（日期3的新航班）
        dep_idx = dep_flights[j]['index']
        combined_dep_df.at[dep_idx, 'ID'] = pair_id
        combined_dep_df.at[dep_idx, '机型'] = k

        pair_id += 1
        pair_num += 1

filtered_dep_df = combined_dep_df[
    (combined_dep_df['日期'] != 3) |  # 保留所有日期非3的航班
    (combined_dep_df['ID'].notna())   # 保留日期3且已配对的航班
    ]

# 保存结果
with pd.ExcelWriter('final_pairing.xlsx') as writer:
    arr_df.to_excel(writer, sheet_name='到达航班', index=False)
    filtered_dep_df.to_excel(writer, sheet_name='出发航班', index=False)

print("需配对航班量: ",len(unpaired_arr))
print("配对成功航班量: ",pair_num)
if len(unpaired_arr) > pair_num:
    print(f"有 {len(unpaired_arr)-pair_num} 个航班未完成配对，请检查！！！")