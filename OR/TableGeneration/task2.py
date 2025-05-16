from collections import defaultdict

import pandas as pd
from gurobipy import Model, GRB, quicksum

from utils import quotas, min_times, dep_hour_distribution, DOM_INT_MIX_p
from utils2 import peak_configs

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
mixed_market_vars = []  # 存储混接配对变量
for i, arr in enumerate(arr_flights):
    for j, dep in enumerate(dep_flights):
        # 新的配对条件判断
        valid_pair = False
        # 条件1：原有同市场配对
        if arr['market'] == dep['market'] and (
                (arr['date'] == dep['date']) or
                (arr['date'] == 1 and dep['date'] == 2)
        ):
            valid_pair = True
        # 条件2：混接配对（日期均为2）
        elif arr['date'] == 2 and dep['date'] == 2:
            valid_pair = True

        if not valid_pair:
            continue

        # 统一计算过站时间（处理所有可能的跨天情况）
        delta = 0
        if arr['date'] == 2 and dep['date'] == 2:
            delta = dep['time'] - arr['time']
            if delta <= 0:
                continue
        elif arr['date'] == 1 and dep['date'] == 2:
            delta = (24 * 60 - arr['time']) + dep['time']

        # 确定适用的市场和机型要求
        if arr['market'] == dep['market']:  # 同市场
            market = arr['market']
            min_time_config = min_times.get((market, 'C'), 0)  # 默认C类
        else:  # 混接市场
            market_arr = arr['market']
            market_dep = dep['market']
            # 取两个市场机型要求的最大值
            min_time_config = max(
                min_times.get((market_arr, 'C'), 0),
                min_times.get((market_dep, 'C'), 0)
            )

        # 生成机型变量
        for k in ['C', 'E', 'F']:
            min_time_needed = min_times.get((arr['market'], k), 0)
            if delta < min_time_needed:
                continue

            # 配额检查
            arr_quota = quotas.get((arr['market'], k), {}).get('ARR', 0)
            dep_quota = quotas.get((dep['market'], k), {}).get('DEP', 0)

            if arr_quota > 0 and dep_quota > 0:
                var = model.addVar(vtype=GRB.BINARY, name=f'x_{i}_{j}_{k}')
                variables[(i, j, k)] = var
                # 记录混接配对
                if arr['market'] != dep['market']:
                    mixed_market_vars.append(var)

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
                   (arr_flights[arr_idx]['date'] == 1)  # 跨日期配对
           )
    ]

    if not possible_pairs:
        raise ValueError(f"日期2的出发航班{j}没有可用的到达航班配对，请检查数据")

    # 添加必须配对的约束
    model.addConstr(
        quicksum(possible_pairs) == 1,
        name=f"mandatory_pairing_date2_dep_{j}"
    )

# 约束6：宽体机型在第二天的离港小时分布限制
# 创建小时限制字典
hour_limits = {}
for h, dom_dep, int_dep in dep_hour_distribution:
    hour_limits[(h, 'DOM')] = dom_dep
    hour_limits[(h, 'INT')] = int_dep

# 按小时和市场分组宽体机型离港变量
hourly_dep_vars = defaultdict(list)
for (i, j, k), var in variables.items():
    # 仅处理宽体机型且日期为2的出发航班
    if k in ['E', 'F'] and dep_flights[j]['date'] == 2:
        dep_time = dep_flights[j]['time']
        dep_hour = dep_time // 60  # 转换为小时
        market = dep_flights[j]['market']
        hourly_dep_vars[(dep_hour, market)].append(var)

# 添加小时分布约束
for (h, market), vars_list in hourly_dep_vars.items():
    if (h, market) in hour_limits:
        # 检查hour_limits中的值是否为nan
        if pd.isna(hour_limits[(h, market)]):
            continue
        original = hour_limits[(h, market)]
        # 计算允许的上下限
        bias = 0
        upper_bound = original + bias
        lower_bound = max(0, original - bias)  # 保证下限不低于0

        # 添加柔性约束
        model.addConstr(
            quicksum(vars_list) <= upper_bound,
            name=f"flex_hourly_upper_{market}_h{h}"
        )
        # 仅当原限制>0时添加下限约束
        if original > 0:
            model.addConstr(
                quicksum(vars_list) >= lower_bound,
                name=f"flex_hourly_lower_{market}_h{h}"
            )

# 约束7：国内国际高峰小时机型比例约束
for config in peak_configs:
    # 转换时间到分钟
    start_h, start_m = map(int, config['start_time'].split(':'))
    end_h, end_m = map(int, config['end_time'].split(':'))
    start_min = start_h * 60 + start_m
    end_min = end_h * 60 + end_m

    # 计算机型配额
    ratios = config['ratios']
    e_count = round(config['total'] * ratios.get('E', 0))
    f_count = round(config['total'] * ratios.get('F', 0))
    c_count = config['total'] - e_count - f_count

    # 收集变量
    c_vars = []
    e_vars = []
    f_vars = []

    # 根据方向筛选航班
    direction = config['direction']
    market = config['market']

    if direction == 'ARR':
        # 到达航班筛选
        selected = [
            i for i, arr in enumerate(arr_flights)
            if arr['market'] == market
               and arr['date'] == 2
               and start_min <= arr['time'] <= end_min
        ]
        # 收集到达航班对应的机型变量
        for (i, j, k), var in variables.items():
            if i in selected:
                if k == 'C':
                    c_vars.append(var)
                elif k == 'E':
                    e_vars.append(var)
                elif k == 'F':
                    f_vars.append(var)

    elif direction == 'DEP':
        # 出发航班筛选
        selected = [
            j for j, dep in enumerate(dep_flights)
            if dep['market'] == market
               and dep['date'] == 2
               and start_min <= dep['time'] <= end_min
        ]
        # 收集出发航班对应的机型变量
        for (i, j, k), var in variables.items():
            if j in selected:
                if k == 'C':
                    c_vars.append(var)
                elif k == 'E':
                    e_vars.append(var)
                elif k == 'F':
                    f_vars.append(var)

    elif direction == 'both':
        # 双向筛选（同时考虑到达和出发）
        selected_arr = [
            i for i, arr in enumerate(arr_flights)
            if arr['market'] == market
               and arr['date'] == 2
               and start_min <= arr['time'] <= end_min
        ]
        selected_dep = [
            j for j, dep in enumerate(dep_flights)
            if dep['market'] == market
               and dep['date'] == 2
               and start_min <= dep['time'] <= end_min
        ]
        # 收集双向变量
        for (i, j, k), var in variables.items():
            if i in selected_arr or j in selected_dep:
                if k == 'C':
                    c_vars.append(var)
                elif k == 'E':
                    e_vars.append(var)
                elif k == 'F':
                    f_vars.append(var)

    # 添加约束
    if c_count >= 0 and c_vars:
        model.addConstr(quicksum(c_vars) <= c_count, name=f"peak_{config['name']}_C")
    if e_count >= 0 and e_vars:
        model.addConstr(quicksum(e_vars) <= e_count, name=f"peak_{config['name']}_E")
    if f_count >= 0 and f_vars:
        model.addConstr(quicksum(f_vars) <= f_count, name=f"peak_{config['name']}_F")

# 约束8：国内国际混接比例：5% （仅讨论第二天天内）
if mixed_market_vars:
    total_pairs = quicksum(variables.values())
    model.addConstr(
        quicksum(mixed_market_vars) <= DOM_INT_MIX_p * total_pairs,
        name="mixed_market_limit"
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
