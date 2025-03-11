import pandas as pd
from gurobipy import Model, GRB, quicksum

# 机型和市场配额约束（日期均为2）
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

# 读取Excel文件中的到达和出发航班数据
arr_df = pd.read_excel('pre_task2.xlsx', sheet_name='到达航班')
dep_df = pd.read_excel('pre_task2.xlsx', sheet_name='出发航班')


# 预处理函数：将时间字符串转换为分钟数
def time_to_minutes(t_str):
    h, m = map(int, t_str.split(':'))
    return h * 60 + m


# 解析到达和出发航班数据
arr_flights = []
for idx, row in arr_df.iterrows():
    arr_flights.append({
        'index': idx,
        'time': time_to_minutes(row['时间']),
        'market': row['市场'],
        'date': row['日期'],
        'direction': 'ARR'
    })

dep_flights = []
for idx, row in dep_df.iterrows():
    dep_flights.append({
        'index': idx,
        'time': time_to_minutes(row['时间']),
        'market': row['市场'],
        'date': row['日期'],
        'direction': 'DEP'
    })

# 初始化模型
model = Model('Flight_Pairing')

# 生成变量：x[(i, j, k)]表示到达航班i与出发航班j配对且机型为k
variables = {}
for i, arr in enumerate(arr_flights):
    for j, dep in enumerate(dep_flights):
        # 允许日期为1的到达与日期为2的出发配对
        if arr['market'] == dep['market'] and (
                (arr['date'] == dep['date']) or  # 同日期配对
                (arr['date'] == 1 and dep['date'] == 2)  # 跨日期配对
        ):
            # 计算实际过站时间（处理跨天情况）
            if arr['date'] == dep['date']:
                delta = dep['time'] - arr['time']
            else:  # 日期1到达 + 日期2出发
                delta = (24 * 60 - arr['time']) + dep['time']

            # 排除无效时间差（即使跨天也要保证delta>0）
            if delta <= 0:
                continue

            market = arr['market']
            # 检查所有可能的机型
            for k in ['C', 'E', 'F']:
                if (market, k) in min_times and delta >= min_times[(market, k)]:
                    if quotas[(market, k)]['ARR'] > 0 and quotas[(market, k)]['DEP'] > 0:
                        variables[(i, j, k)] = model.addVar(vtype=GRB.BINARY, name=f'x_{i}_{j}_{k}')

# 目标函数：最大化配对总数
model.setObjective(quicksum(variables.values()), GRB.MAXIMIZE)

# 约束1：每个到达航班最多配对一次
arr_used = {}
for key in variables:
    i, j, k = key
    if i not in arr_used:
        arr_used[i] = []
    arr_used[i].append(variables[key])
for i in arr_used:
    model.addConstr(quicksum(arr_used[i]) <= 1, name=f'arr_{i}')

# 约束2：每个出发航班最多配对一次
dep_used = {}
for key in variables:
    i, j, k = key
    if j not in dep_used:
        dep_used[j] = []
    dep_used[j].append(variables[key])
for j in dep_used:
    model.addConstr(quicksum(dep_used[j]) <= 1, name=f'dep_{j}')

# 约束3：机型配额限制
quota_constraints = {}
for (market, k), quota in quotas.items():
    arr_total = []
    dep_total = []
    for key in variables:
        # 正确解包元组的三个元素
        i, j, key_k = key
        # 过滤当前市场机型
        if key_k == k and arr_flights[i]['market'] == market and arr_flights[i]['date'] == 2:
            arr_total.append(variables[key])
        if key_k == k and dep_flights[j]['market'] == market and dep_flights[j]['date'] == 2:
            dep_total.append(variables[key])
    if arr_total:
        model.addConstr(quicksum(arr_total) <= quota['ARR'], f'arr_quota_{market}_{k}')
    if dep_total:
        model.addConstr(quicksum(dep_total) <= quota['DEP'], f'dep_quota_{market}_{k}')


# 约束4：日期1的到达航班必须全部配对
date1_arr_indices = [i for i, arr in enumerate(arr_flights) if arr['date'] == 1]

for i in date1_arr_indices:
    # 找到该航班所有可能的配对变量
    possible_pairs = [
        var for (arr_idx, dep_idx, k), var in variables.items()
        if arr_idx == i  # 匹配当前到达航班
           and dep_flights[dep_idx]['date'] == 2  # 确保是第二天的出发
    ]

    if not possible_pairs:
        raise ValueError(f"日期1的到达航班{i}没有可用的出发航班配对，请检查数据")

    # 添加必须配对的约束
    model.addConstr(
        quicksum(possible_pairs) == 1,
        name=f"mandatory_pairing_date1_arr_{i}"
    )


# 约束5：日期2的所有出发航班必须配对
date2_dep_indices = [j for j, dep in enumerate(dep_flights) if dep['date'] == 2]

for j in date2_dep_indices:
    # 找到该出发航班所有可能的配对变量
    possible_pairs = [
        var for (arr_idx, dep_idx, k), var in variables.items()
        if dep_idx == j  # 匹配当前出发航班
           and (
                   (arr_flights[arr_idx]['date'] == 2) or  # 同日期配对
                   (arr_flights[arr_idx]['date'] == 1)     # 跨日期配对
           )
    ]

    if not possible_pairs:
        raise ValueError(f"日期2的出发航班{j}没有可用的到达航班配对，请检查数据")

    # 添加必须配对的约束
    model.addConstr(
        quicksum(possible_pairs) == 1,
        name=f"mandatory_pairing_date2_dep_{j}"
    )

model.optimize()

# 收集配对结果
pair_id = 1
pair_results = []
for key in variables:
    if variables[key].X > 0.5:
        i, j, k = key
        pair_results.append({
            'arr_idx': arr_flights[i]['index'],
            'dep_idx': dep_flights[j]['index'],
            'pair_id': pair_id,
            'aircraft_type': k
        })
        pair_id += 1

# 将结果写入DataFrame
for pair in pair_results:
    arr_df.at[pair['arr_idx'], 'ID'] = pair['pair_id']
    arr_df.at[pair['arr_idx'], '机型'] = pair['aircraft_type']
    dep_df.at[pair['dep_idx'], 'ID'] = pair['pair_id']
    dep_df.at[pair['dep_idx'], '机型'] = pair['aircraft_type']

# 保存到Excel
with pd.ExcelWriter('output_flight_pairing.xlsx') as writer:
    arr_df.to_excel(writer, sheet_name='到达航班', index=False)
    dep_df.to_excel(writer, sheet_name='出发航班', index=False)
