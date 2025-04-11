import gurobipy as gp
import pandas as pd
from gurobipy import GRB

from dataset import DOM_AIRLINES, INT_AIRLINES, HEADINGS_ARR, HEADINGS_DEP

# 文件路径
CURRENT_STATUS_FILE = "2024_现状.xlsx"
PAIRING_FILE = "final_pairing_processed.xlsx"
OUTPUT_FILE = "final_pairing_processed_with_assignments.xlsx"
ARR_HOURLY_STATS_FILE = "机_到达航班统计表.xlsx"  # 新增文件路径
DEP_HOURLY_STATS_FILE = "出发航班比例表.xlsx"


def load_data():
    """加载所有需要的数据"""
    # 加载现状数据
    current_status = pd.read_excel(CURRENT_STATUS_FILE)

    # 加载配对数据
    arrival_flights = pd.read_excel(PAIRING_FILE, sheet_name="到达航班")
    departure_flights = pd.read_excel(PAIRING_FILE, sheet_name="出发航班")

    # 加载机场到达航班统计表
    hourly_dom_stats = pd.read_excel(ARR_HOURLY_STATS_FILE, sheet_name="国内航班")
    hourly_int_stats = pd.read_excel(ARR_HOURLY_STATS_FILE, sheet_name="国际航班")

    # 加载出发航班统计表
    hourly_dom_dep_stats = pd.read_excel(DEP_HOURLY_STATS_FILE, sheet_name="国内航班")
    hourly_int_dep_stats = pd.read_excel(DEP_HOURLY_STATS_FILE, sheet_name="国际航班")

    return current_status, arrival_flights, departure_flights, hourly_dom_stats, hourly_int_stats, hourly_dom_dep_stats, hourly_int_dep_stats


def get_hour_from_time(time_str):
    """从时间字符串中提取小时"""
    if pd.isna(time_str):
        return None

    if isinstance(time_str, str):
        parts = time_str.split(':')
        if len(parts) >= 2:
            return int(parts[0])
    elif isinstance(time_str, (int, float)):
        # 假设时间格式为小数，如 23.55 表示 23:55
        return int(time_str)

    return None


def get_main_headings(current_status):
    """
    从现状数据中获取各航司的主航向
    
    Args:
        current_status: 现状数据
    
    Returns:
        包含各航司主航向的字典，格式为 {(航司, 市场类型, 方向): (主航向, 比例)}
    """
    main_headings = {}

    # 处理进港航班
    arr_data = current_status[current_status['Direction'] == 'Arrival']

    # 处理国内进港
    dom_arr = arr_data[arr_data['flightnature'] == 'DOM']
    dom_arr_counts = dom_arr.groupby(['AirlineGroup', 'Routing']).size().reset_index(name='count')

    # 处理国际进港
    int_arr = arr_data[arr_data['flightnature'] == 'INT']
    int_arr_counts = int_arr.groupby(['AirlineGroup', 'Routing']).size().reset_index(name='count')

    # 处理离港航班
    dep_data = current_status[current_status['Direction'] == 'Departure']

    # 处理国内离港
    dom_dep = dep_data[dep_data['flightnature'] == 'DOM']
    dom_dep_counts = dom_dep.groupby(['AirlineGroup', 'Routing']).size().reset_index(name='count')

    # 处理国际离港
    int_dep = dep_data[dep_data['flightnature'] == 'INT']
    int_dep_counts = int_dep.groupby(['AirlineGroup', 'Routing']).size().reset_index(name='count')

    # 获取各航司的主航向
    for airline in set(current_status['AirlineGroup']):
        if airline == "其他":
            continue

        # 国内进港主航向
        airline_dom_arr = dom_arr_counts[dom_arr_counts['AirlineGroup'] == airline]
        if not airline_dom_arr.empty:
            total_flights = airline_dom_arr['count'].sum()
            main_heading = airline_dom_arr.loc[airline_dom_arr['count'].idxmax()]
            main_heading_ratio = main_heading['count'] / total_flights
            main_headings[(airline, 'DOM', 'Arrival')] = (main_heading['Routing'], main_heading_ratio)

        # 国际进港主航向
        airline_int_arr = int_arr_counts[int_arr_counts['AirlineGroup'] == airline]
        if not airline_int_arr.empty:
            total_flights = airline_int_arr['count'].sum()
            main_heading = airline_int_arr.loc[airline_int_arr['count'].idxmax()]
            main_heading_ratio = main_heading['count'] / total_flights
            main_headings[(airline, 'INT', 'Arrival')] = (main_heading['Routing'], main_heading_ratio)

        # 国内离港主航向
        airline_dom_dep = dom_dep_counts[dom_dep_counts['AirlineGroup'] == airline]
        if not airline_dom_dep.empty:
            total_flights = airline_dom_dep['count'].sum()
            main_heading = airline_dom_dep.loc[airline_dom_dep['count'].idxmax()]
            main_heading_ratio = main_heading['count'] / total_flights
            main_headings[(airline, 'DOM', 'Departure')] = (main_heading['Routing'], main_heading_ratio)

        # 国际离港主航向
        airline_int_dep = int_dep_counts[int_dep_counts['AirlineGroup'] == airline]
        if not airline_int_dep.empty:
            total_flights = airline_int_dep['count'].sum()
            main_heading = airline_int_dep.loc[airline_int_dep['count'].idxmax()]
            main_heading_ratio = main_heading['count'] / total_flights
            main_headings[(airline, 'INT', 'Departure')] = (main_heading['Routing'], main_heading_ratio)

    return main_headings


def assign_arrival_flights(current_status, arrival_flights, market_type, hourly_dom_stats, hourly_int_stats,
                           departure_flights=None, prev_dom_dep_counts=None, prev_int_dep_counts=None):
    """
    为进港航班分配航司和航向

    Args:
        current_status: 现状数据
        arrival_flights: 需要分配的进港航班
        market_type: 'DOM' 或 'INT'
        hourly_dom_stats: 国内航班小时统计数据
        hourly_int_stats: 国际航班小时统计数据
        departure_flights: 所有离港航班数据，用于检查配额限制
        prev_dom_dep_counts: 之前分配的国内离港航班数量
        prev_int_dep_counts: 之前分配的国际离港航班数量

    Returns:
        带有航司和航向分配的进港航班数据，以及各航司的DOM和INT离港航班分配数量
    """
    print(f"开始分配{market_type}进港航班...")

    main_headings = get_main_headings(current_status)

    # 选择对应的小时统计数据
    hourly_stats = hourly_dom_stats if market_type == 'DOM' else hourly_int_stats
    # 筛选出日期为2的进港航班
    day2_arrivals = arrival_flights[(arrival_flights['日期'] == 2) & (arrival_flights['市场'] == market_type)]

    # flight_mapping = {
    #     row['ID']: (row['机型'], row['小时'])
    #     for _, row in day2_arrivals.iterrows()
    # }

    # 筛选现状数据中的进港航班
    current_arrivals = current_status[
        (current_status['Direction'] == 'Arrival') &
        (current_status['flightnature'] == market_type)
        ]

    # 为每个航班添加小时列
    day2_arrivals['小时'] = day2_arrivals['时间'].apply(get_hour_from_time)
    current_arrivals['小时'] = current_arrivals['time'].apply(get_hour_from_time)

    # 统计现状中每个小时、每个航司、每个航向的航班数量
    current_hour_airline_count = current_arrivals.groupby(['小时', 'AirlineGroup']).size().reset_index(name='count')
    current_hour_airline_heading_count = current_arrivals.groupby(
        ['小时', 'AirlineGroup', 'Routing']).size().reset_index(name='count')
    current_hour_airline_heading_acft_count = current_arrivals.groupby(
        ['小时', 'AirlineGroup', 'Routing', 'Acft Cat']).size().reset_index(name='count')

    # 获取所有可能的航司和航向
    if market_type == 'DOM':
        airlines_data = DOM_AIRLINES
        headings_data = HEADINGS_ARR
    else:
        airlines_data = INT_AIRLINES
        headings_data = HEADINGS_ARR

    all_airlines = []
    for base_type in airlines_data:
        for airline in airlines_data[base_type]:
            all_airlines.append((airline, base_type))

    all_headings = []
    for heading in headings_data:
        heading_type = headings_data[heading]['INT性质'] if market_type == 'INT' else None
        all_headings.append((heading, heading_type))

    # 创建Gurobi模型
    model = gp.Model("Arrival_Assignment")

    # 创建决策变量
    # x[i, a, h] = 1 表示航班i分配给航司a和航向h
    x = {}
    for i, flight in day2_arrivals.iterrows():
        flight_id = flight['ID']
        flight_hour = flight['小时']
        flight_acft = flight['机型']

        for airline, base_type in all_airlines:
            for heading, heading_type in all_headings:
                # 检查是否满足约束条件
                # 1. 绝对远程航向只分配给主基地航司或东航南航海航集团
                if market_type == 'INT' and heading_type == '绝对远程' and base_type == '非主基地' and airline not in [
                    '东航集团', '南航集团', '海航集团']:
                    continue

                # 2. 绝对远程航向分配宽体机
                if market_type == 'INT' and heading_type == '绝对远程' and flight_acft not in ['E', 'F']:
                    continue

                x[flight_id, airline, heading] = model.addVar(vtype=GRB.BINARY,
                                                              name=f"x_{flight_id}_{airline}_{heading}")

    # 每个航班只能分配给一个航司和一个航向
    for i, flight in day2_arrivals.iterrows():
        flight_id = flight['ID']
        model.addConstr(
            gp.quicksum(x[flight_id, airline, heading]
                        for airline, _ in all_airlines
                        for heading, _ in all_headings
                        if (flight_id, airline, heading) in x) == 1,
            f"one_assignment_{flight_id}"
        )

    # 航司配额约束
    for airline, base_type in all_airlines:
        if market_type == 'DOM':
            quota = DOM_AIRLINES[base_type][airline]['ARR']
        else:
            quota = INT_AIRLINES[base_type][airline]['ARR']

        model.addConstr(
            gp.quicksum(x[flight_id, airline, heading]
                        for i, flight in day2_arrivals.iterrows()
                        for heading, _ in all_headings
                        if (flight_id := flight['ID'], airline, heading) in x) == quota,
            f"airline_quota_{airline}"
        )

    # 航向配额约束
    for heading, heading_type in all_headings:
        if market_type == 'DOM':
            quota = HEADINGS_ARR[heading]['DOM']
        else:
            quota = HEADINGS_ARR[heading]['INT']

        model.addConstr(
            gp.quicksum(x[flight_id, airline, heading]
                        for i, flight in day2_arrivals.iterrows()
                        for airline, _ in all_airlines
                        if (flight_id := flight['ID'], airline, heading) in x) == quota,
            f"heading_quota_{heading}"
        )

    # 主航向比例约束：确保各航司的主航向比例不低于现状
    for airline, _ in all_airlines:
        if airline == "其他":
            continue

        key = (airline, market_type, 'Arrival')
        if key in main_headings:
            main_heading, current_ratio = main_headings[key]

            # 计算该航司的总航班数
            total_flights_expr = gp.quicksum(x[flight_id, airline, heading]
                                             for i, flight in day2_arrivals.iterrows()
                                             for heading, _ in all_headings
                                             if (flight_id := flight['ID'], airline, heading) in x)

            # 计算该航司分配给主航向的航班数
            main_heading_flights_expr = gp.quicksum(x[flight_id, airline, main_heading]
                                                    for i, flight in day2_arrivals.iterrows()
                                                    if (flight_id := flight['ID'], airline, main_heading) in x)

            # 添加比例约束
            if total_flights_expr.size() > 0:  # 确保有航班分配给该航司
                model.addConstr(
                    main_heading_flights_expr >= current_ratio * total_flights_expr,
                    f"main_heading_ratio_{airline}_{market_type}_Arrival"
                )

                # 添加新约束：确保主航向仍然是航班量最多的航向
                for heading, _ in all_headings:
                    if heading != main_heading:
                        heading_flights_expr = gp.quicksum(x[flight_id, airline, heading]
                                                           for i, flight in day2_arrivals.iterrows()
                                                           if (flight_id := flight['ID'], airline, heading) in x)

                        model.addConstr(
                            main_heading_flights_expr >= heading_flights_expr,
                            f"main_heading_dominance_{airline}_{market_type}_{heading}_Arrival"
                        )

    dom_expressions = {}  # 存储各航司的DOM表达式
    int_expressions = {}

    # 添加新约束：确保每个航司分配的进港航班对应的离港航班总量不超过配额
    if departure_flights is not None:
        # 筛选出日期为2的离港航班
        day2_departures = departure_flights[departure_flights['日期'] == 2]

        # 获取进港航班ID对应的离港航班
        arrival_to_departure = {}
        for i, flight in day2_arrivals.iterrows():
            flight_id = flight['ID']
            # 查找相同ID的离港航班
            matching_departures = day2_departures[day2_departures['ID'] == flight_id]
            if not matching_departures.empty:
                arrival_to_departure[flight_id] = matching_departures.iloc[0]['市场']

        # 为每个航司添加离港航班配额约束
        for airline, base_type in all_airlines:
            # 获取该航司的国内和国际离港配额
            dom_quota = DOM_AIRLINES[base_type][airline]['DEP']
            int_quota = INT_AIRLINES[base_type][airline]['DEP']

            # 计算分配给该航司的进港航班中，对应的国内和国际离港航班数量
            dom_dep_expr = gp.quicksum(x[flight_id, airline, heading]
                                       for i, flight in day2_arrivals.iterrows()
                                       for heading, _ in all_headings
                                       if (flight_id := flight['ID']) in arrival_to_departure
                                       and arrival_to_departure[flight_id] == 'DOM'
                                       and (flight_id, airline, heading) in x)

            int_dep_expr = gp.quicksum(x[flight_id, airline, heading]
                                       for i, flight in day2_arrivals.iterrows()
                                       for heading, _ in all_headings
                                       if (flight_id := flight['ID']) in arrival_to_departure
                                       and arrival_to_departure[flight_id] == 'INT'  
                                       and (flight_id, airline, heading) in x)

            dom_expressions[airline] = dom_dep_expr
            int_expressions[airline] = int_dep_expr

            # 考虑之前分配的航班数量
            prev_dom_count = prev_dom_dep_counts.get(airline, 0) if prev_dom_dep_counts else 0
            prev_int_count = prev_int_dep_counts.get(airline, 0) if prev_int_dep_counts else 0

            # 添加约束：对应的国内离港航班数量不超过配额
            model.addConstr(dom_dep_expr + prev_dom_count <= dom_quota, f"dom_dep_quota_{airline}")

            # 添加约束：对应的国际离港航班数量不超过配额
            model.addConstr(int_dep_expr + prev_int_count <= int_quota, f"int_dep_quota_{airline}")


    # 添加基于小时统计数据的约束和目标函数
    # 目标函数：最大化与小时统计数据的匹配度
    obj_expr = gp.LinExpr()

    # 首先，计算当前分布的小时分布情况
    current_hour_distribution = {}
    for _, row in hourly_stats.iterrows():
        airline = row['AirlineGroup']
        routing = row['Routing']
        acft_cat = row['Acft Cat']

        # 跳过不在我们考虑范围内的航司
        if airline == '其它':
            continue

        # 获取该行的小时分布数据
        for hour in range(24):
            hour_col = hour
            if hour_col in row and not pd.isna(row[hour_col]):
                key = (airline, routing, acft_cat, hour)
                current_hour_distribution[key] = row[hour_col]

    # 创建变量来表示未来分布
    future_distribution = {}
    for hour in range(24):
        hour_flights = day2_arrivals[day2_arrivals['小时'] == hour]

        for airline, _ in all_airlines:
            for heading, _ in all_headings:
                for acft_cat in ['C', 'E']:  # 简化为窄体机和宽体机两类
                    key = (airline, heading, acft_cat, hour)

                    # 计算该组合的航班数量表达式
                    count_expr = gp.LinExpr()

                    for i, flight in hour_flights.iterrows():
                        flight_id = flight['ID']
                        flight_acft = flight['机型']

                        # 检查机型是否匹配
                        is_matching_acft = (acft_cat == 'C' and flight_acft in ['A', 'B', 'C', 'D']) or \
                                           (acft_cat == 'E' and flight_acft in ['E', 'F'])

                        if is_matching_acft and (flight_id, airline, heading) in x:
                            count_expr += x[flight_id, airline, heading]

                    # 如果有航班可能被分配到这个组合
                    if count_expr.size() > 0:
                        # 创建变量表示该组合的航班数量
                        var_name = f"future_{airline}_{heading}_{acft_cat}_{hour}"
                        future_distribution[key] = model.addVar(name=var_name, lb=0, vtype=GRB.INTEGER)

                        # 添加约束，确保future_distribution变量等于count_expr
                        model.addConstr(future_distribution[key] == count_expr,
                                        f"future_distr_{airline}_{heading}_{acft_cat}_{hour}")

    # 计算总航班数
    total_flights = day2_arrivals.shape[0]

    # 计算未来分布与当前分布的偏离度
    for key in set(current_hour_distribution.keys()) | set(future_distribution.keys()):
        current_value = current_hour_distribution.get(key, 0)

        if key in future_distribution:
            # 如果未来分布中有这个组合，计算偏离度
            future_value = future_distribution[key] / total_flights  # 转换为比例

            # 添加偏离度到目标函数（使用平方差）
            deviation_var = model.addVar(name=f"dev_{key[0]}_{key[1]}_{key[2]}_{key[3]}", lb=0)

            # 添加约束来定义偏离度变量
            model.addConstr(deviation_var >= future_value - current_value,
                            f"dev_pos_{key[0]}_{key[1]}_{key[2]}_{key[3]}")
            model.addConstr(deviation_var >= current_value - future_value,
                            f"dev_neg_{key[0]}_{key[1]}_{key[2]}_{key[3]}")

            # 将偏离度添加到目标函数，权重可以根据需要调整
            obj_expr += deviation_var * 100
        else:
            # 如果未来分布中没有这个组合，但当前分布有，则偏离度为当前值
            if current_value > 0:
                obj_expr += current_value * 100

    # 设置目标函数为最小化偏离度
    model.setObjective(obj_expr, GRB.MINIMIZE)

    # 求解模型
    model.optimize()
    # 创建用于返回的离港航班分配数量字典

    dom_dep_counts = {}
    int_dep_counts = {}
    if model.status == gp.GRB.OPTIMAL:
        for airline, expr in dom_expressions.items():
            print(f"{airline} DOM:", expr.getValue())
        for airline, expr in int_expressions.items():
            print(f"{airline} INT:", expr.getValue())

    if model.status == gp.GRB.OPTIMAL:
        # 计算各航司的离港航班分配数量
        for airline in dom_expressions:
            dom_dep_counts[airline] = dom_expressions[airline].getValue()
        for airline in int_expressions:
            int_dep_counts[airline] = int_expressions[airline].getValue()

        print(f"{market_type}进港航班分配成功！")

        # 将结果添加到航班数据中
        result_df = day2_arrivals.copy()
        result_df['航司'] = ""
        result_df['航向'] = ""

        for i, flight in day2_arrivals.iterrows():
            flight_id = flight['ID']
            for airline, _ in all_airlines:
                for heading, _ in all_headings:
                    if (flight_id, airline, heading) in x and x[flight_id, airline, heading].X > 0.5:
                        result_df.loc[i, '航司'] = airline
                        result_df.loc[i, '航向'] = heading

        return result_df, dom_dep_counts, int_dep_counts
    else:
        print(f"{market_type}进港航班分配失败！模型状态：{model.status}")
        return None, {}, {}

def assign_departure_flights(arrival_assignments, departure_flights, current_status, market_type, hourly_dom_dep_stats,
                             hourly_int_dep_stats):
    """
    为离港航班分配航司和航向
    
    Args:
        arrival_assignments: 已分配的进港航班
        departure_flights: 需要分配的离港航班
        current_status: 现状数据
        market_type: 'DOM' 或 'INT'
        hourly_dom_dep_stats: 国内出发航班小时统计数据
        hourly_int_dep_stats: 国际出发航班小时统计数据
    
    Returns:
        带有航司和航向分配的离港航班数据
    """
    print(f"开始分配{market_type}离港航班...")
    main_headings = get_main_headings(current_status)

    # 选择对应的小时统计数据
    hourly_stats = hourly_dom_dep_stats if market_type == 'DOM' else hourly_int_dep_stats

    # 筛选出日期为2的离港航班
    day2_departures = departure_flights[(departure_flights['日期'] == 2) & (departure_flights['市场'] == market_type)]

    # 筛选现状数据中的离港航班
    current_departures = current_status[
        (current_status['Direction'] == 'Departure') &
        (current_status['flightnature'] == market_type)
        ]

    # 为每个航班添加小时列
    day2_departures['小时'] = day2_departures['时间'].apply(get_hour_from_time)
    current_departures['小时'] = current_departures['time'].apply(get_hour_from_time)

    # 统计现状中每个小时、每个航司、每个航向的航班数量
    current_hour_airline_heading_count = current_departures.groupby(
        ['小时', 'AirlineGroup', 'Routing']).size().reset_index(name='count')
    current_hour_airline_heading_acft_count = current_departures.groupby(
        ['小时', 'AirlineGroup', 'Routing', 'Acft Cat']).size().reset_index(name='count')

    # 获取所有可能的航司和航向
    if market_type == 'DOM':
        airlines_data = DOM_AIRLINES
        headings_data = HEADINGS_DEP
    else:
        airlines_data = INT_AIRLINES
        headings_data = HEADINGS_DEP

    all_airlines = []
    for base_type in airlines_data:
        for airline in airlines_data[base_type]:
            all_airlines.append((airline, base_type))

    all_headings = []
    for heading in headings_data:
        heading_type = headings_data[heading]['INT性质'] if market_type == 'INT' else None
        all_headings.append((heading, heading_type))

    # 创建结果DataFrame
    result_df = day2_departures.copy()
    result_df['航司'] = ""
    result_df['航向'] = ""

    # 为每个离港航班分配航司（与相同ID的进港航班相同）
    # 修改：添加特例处理，检查配对的进港航班日期
    for i, flight in day2_departures.iterrows():
        flight_id = flight['ID']
        # 查找相同ID的进港航班
        matching_arrivals = arrival_assignments[arrival_assignments['ID'] == flight_id]

        if not matching_arrivals.empty:
            # 如果进港航班日期为2，则遵循"离港航司=进港航司"的规则
            # 如果进港航班日期为1，则不需要遵循该规则，航司将在后续步骤中分配
            result_df.loc[i, '航司'] = matching_arrivals.iloc[0]['航司']

    # 创建Gurobi模型来分配航向和未分配航司的航班
    model = gp.Model("Departure_Assignment")

    # 创建决策变量
    # y[i, h] = 1 表示航班i分配给航向h
    y = {}
    # z[i, a] = 1 表示航班i分配给航司a（仅对未分配航司的航班）
    z = {}

    for i, flight in day2_departures.iterrows():
        flight_id = flight['ID']
        flight_hour = flight['小时']
        flight_acft = flight['机型']
        flight_airline = result_df.loc[i, '航司']

        # 如果已经分配了航司
        if flight_airline:
            airline_base_type = next((base_type for airline, base_type in all_airlines if airline == flight_airline),
                                     None)

            for heading, heading_type in all_headings:
                # 检查是否满足约束条件
                # 1. 绝对远程航向只分配给主基地航司或东航南航海航集团
                if market_type == 'INT' and heading_type == '绝对远程' and airline_base_type == '非主基地' and flight_airline not in [
                    '东航集团', '南航集团', '海航集团']:
                    continue

                # 2. 绝对远程航向分配宽体机
                if market_type == 'INT' and heading_type == '绝对远程' and flight_acft not in ['E', 'F']:
                    continue

                y[flight_id, heading] = model.addVar(vtype=GRB.BINARY, name=f"y_{flight_id}_{heading}")
        # 如果未分配航司（进港航班日期为1的情况）
        else:
            # 为航班分配航司
            for airline, base_type in all_airlines:
                z[flight_id, airline] = model.addVar(vtype=GRB.BINARY, name=f"z_{flight_id}_{airline}")

            # 每个航班只能分配给一个航司
            model.addConstr(
                gp.quicksum(z[flight_id, airline]
                            for airline, _ in all_airlines
                            if (flight_id, airline) in z) == 1,
                f"one_airline_{flight_id}"
            )

            # 为航班分配航向
            for airline, base_type in all_airlines:
                for heading, heading_type in all_headings:
                    # 检查是否满足约束条件
                    # 1. 绝对远程航向只分配给主基地航司或东航南航海航集团
                    if market_type == 'INT' and heading_type == '绝对远程' and base_type == '非主基地' and airline not in [
                        '东航集团', '南航集团', '海航集团']:
                        continue

                    # 2. 绝对远程航向分配宽体机
                    if market_type == 'INT' and heading_type == '绝对远程' and flight_acft not in ['E', 'F']:
                        continue

                    # 创建联合变量 w[i, a, h] = z[i, a] * y[i, h]
                    # 由于无法直接表示非线性约束，我们使用一个新变量和线性约束来表示
                    var_name = f"w_{flight_id}_{airline}_{heading}"
                    w = model.addVar(vtype=GRB.BINARY, name=var_name)

                    # 添加约束确保 w[i, a, h] = 1 当且仅当 z[i, a] = 1 且 y[i, h] = 1
                    # 这需要添加到模型中

                    # 为简化处理，我们直接使用 w 变量代替 y 变量
                    y[flight_id, heading, airline] = w

    # 每个航班只能分配给一个航向
    for i, flight in day2_departures.iterrows():
        flight_id = flight['ID']
        flight_airline = result_df.loc[i, '航司']

        if flight_airline:  # 已分配航司
            model.addConstr(
                gp.quicksum(y[flight_id, heading]
                            for heading, _ in all_headings
                            if (flight_id, heading) in y) == 1,
                f"one_heading_{flight_id}"
            )
        else:  # 未分配航司
            for airline, _ in all_airlines:
                if (flight_id, airline) in z:
                    model.addConstr(
                        gp.quicksum(y[flight_id, heading, airline]
                                    for heading, _ in all_headings
                                    if (flight_id, heading, airline) in y) == z[flight_id, airline],
                        f"one_heading_{flight_id}_{airline}"
                    )

    # 航司配额约束（对于未分配航司的航班）
    for airline, base_type in all_airlines:
        if market_type == 'DOM':
            quota = DOM_AIRLINES[base_type][airline]['DEP']
        else:
            quota = INT_AIRLINES[base_type][airline]['DEP']

        # 计算已分配给该航司的航班数量
        assigned_count = sum(1 for i, flight in day2_departures.iterrows()
                             if result_df.loc[i, '航司'] == airline)

        # 剩余配额
        remaining_quota = quota - assigned_count

        # 如果剩余配额大于0，添加约束
        if remaining_quota > 0:
            model.addConstr(
                gp.quicksum(z[flight_id, airline]
                            for i, flight in day2_departures.iterrows()
                            if (flight_id := flight['ID'], airline) in z) == remaining_quota,
                f"airline_quota_{airline}"
            )
        elif remaining_quota < 0:
            print(f"警告：{airline}的离港航班配额已超出，请检查数据")

    # 航向配额约束
    for heading, heading_type in all_headings:
        if market_type == 'DOM':
            quota = HEADINGS_DEP[heading]['DOM']
        else:
            quota = HEADINGS_DEP[heading]['INT']

        # 对于已分配航司的航班
        assigned_expr = gp.quicksum(y[flight_id, heading]
                                    for i, flight in day2_departures.iterrows()
                                    if result_df.loc[i, '航司'] and (flight_id := flight['ID'], heading) in y)

        # 对于未分配航司的航班
        unassigned_expr = gp.quicksum(y[flight_id, heading, airline]
                                      for i, flight in day2_departures.iterrows()
                                      if not result_df.loc[i, '航司']
                                      for airline, _ in all_airlines
                                      if (flight_id := flight['ID'], heading, airline) in y)

        model.addConstr(assigned_expr + unassigned_expr == quota, f"heading_quota_{heading}")

    # 主航向比例约束：确保各航司的主航向比例不低于现状
    for airline, _ in all_airlines:
        if airline == "其他":
            continue

        key = (airline, market_type, 'Departure')
        if key in main_headings:
            main_heading, current_ratio = main_headings[key]

            # 计算已分配该航司的航班数量
            assigned_flights = [i for i, flight in day2_departures.iterrows()
                                if result_df.loc[i, '航司'] == airline]

            # 如果有航班分配给该航司
            if assigned_flights:
                # 计算已分配该航司的主航向航班数量表达式
                assigned_main_heading_expr = gp.quicksum(y[flight_id, main_heading]
                                                         for i in assigned_flights
                                                         if
                                                         (flight_id := day2_departures.loc[i, 'ID'], main_heading) in y)

                # 计算未分配航司的航班中，分配给该航司的主航向航班数量表达式
                unassigned_main_heading_expr = gp.quicksum(y[flight_id, main_heading, airline]
                                                           for i, flight in day2_departures.iterrows()
                                                           if not result_df.loc[i, '航司']
                                                           and (flight_id := flight['ID'], main_heading, airline) in y)

                # 计算该航司的总航班数表达式
                total_assigned_expr = len(assigned_flights)
                total_unassigned_expr = gp.quicksum(z[flight_id, airline]
                                                    for i, flight in day2_departures.iterrows()
                                                    if not result_df.loc[i, '航司']
                                                    and (flight_id := flight['ID'], airline) in z)

                # 添加比例约束
                model.addConstr(
                    (assigned_main_heading_expr + unassigned_main_heading_expr) >=
                    current_ratio * (total_assigned_expr + total_unassigned_expr),
                    f"main_heading_ratio_{airline}_{market_type}_Departure"
                )

                # 添加新约束：确保主航向仍然是航班量最多的航向
                for heading, _ in all_headings:
                    if heading != main_heading:
                        # 计算该航向的航班数量表达式
                        assigned_heading_expr = gp.quicksum(y[flight_id, heading]
                                                            for i in assigned_flights
                                                            if
                                                            (flight_id := day2_departures.loc[i, 'ID'], heading) in y)

                        unassigned_heading_expr = gp.quicksum(y[flight_id, heading, airline]
                                                              for i, flight in day2_departures.iterrows()
                                                              if not result_df.loc[i, '航司']
                                                              and (flight_id := flight['ID'], heading, airline) in y)

                        model.addConstr(
                            (assigned_main_heading_expr + unassigned_main_heading_expr) >=
                            (assigned_heading_expr + unassigned_heading_expr),
                            f"main_heading_dominance_{airline}_{market_type}_{heading}_Departure"
                        )

    # 添加基于小时统计数据的约束和目标函数
    # 目标函数：最小化未来分布与当前分布的偏离度
    obj_expr = gp.LinExpr()

    # 首先，计算当前分布的小时分布情况
    current_hour_distribution = {}
    for _, row in hourly_stats.iterrows():
        airline = row['AirlineGroup']
        routing = row['Routing']
        acft_cat = row['Acft Cat']

        # 跳过不在我们考虑范围内的航司
        if airline == '其它':
            continue

        # 获取该行的小时分布数据
        for hour in range(24):
            hour_col = hour
            if hour_col in row and not pd.isna(row[hour_col]):
                key = (airline, routing, acft_cat, hour)
                current_hour_distribution[key] = row[hour_col]

    # 创建变量来表示未来分布
    future_distribution = {}
    for hour in range(24):
        hour_flights = day2_departures[day2_departures['小时'] == hour]

        for airline, _ in all_airlines:
            for heading, _ in all_headings:
                for acft_cat in ['C', 'E']:  # 简化为窄体机和宽体机两类
                    key = (airline, heading, acft_cat, hour)

                    # 计算该组合的航班数量表达式
                    count_expr = gp.LinExpr()

                    # 已分配航司的航班
                    for i, flight in hour_flights.iterrows():
                        flight_id = flight['ID']
                        flight_airline = result_df.loc[i, '航司']
                        flight_acft = flight['机型']

                        # 检查机型是否匹配
                        is_matching_acft = (acft_cat == 'C' and flight_acft in ['A', 'B', 'C', 'D']) or \
                                           (acft_cat == 'E' and flight_acft in ['E', 'F'])

                        if is_matching_acft and flight_airline == airline and (flight_id, heading) in y:
                            count_expr += y[flight_id, heading]

                    # 未分配航司的航班
                    for i, flight in hour_flights.iterrows():
                        flight_id = flight['ID']
                        flight_airline = result_df.loc[i, '航司']
                        flight_acft = flight['机型']

                        # 检查机型是否匹配
                        is_matching_acft = (acft_cat == 'C' and flight_acft in ['A', 'B', 'C', 'D']) or \
                                           (acft_cat == 'E' and flight_acft in ['E', 'F'])

                        if is_matching_acft and not flight_airline and (flight_id, airline) in z:
                            for heading_item, _ in all_headings:
                                if (flight_id, heading_item, airline) in y and heading_item == heading:
                                    count_expr += y[flight_id, heading_item, airline]

                    # 如果有航班可能被分配到这个组合
                    if count_expr.size() > 0:
                        # 创建变量表示该组合的航班数量
                        var_name = f"future_{airline}_{heading}_{acft_cat}_{hour}"
                        future_distribution[key] = model.addVar(name=var_name, lb=0, vtype=GRB.INTEGER)

                        # 添加约束，确保future_distribution变量等于count_expr
                        model.addConstr(future_distribution[key] == count_expr,
                                        f"future_distr_{airline}_{heading}_{acft_cat}_{hour}")

                # 对于机型分布
                for acft_cat in ['C', 'E']:  # 简化为窄体机和宽体机两类
                    # 计算该组合的航班数量表达式
                    acft_count_expr = gp.LinExpr()

                    # 已分配航司的航班
                    for i, flight in hour_flights.iterrows():
                        flight_id = flight['ID']
                        flight_airline = result_df.loc[i, '航司']
                        flight_acft = flight['机型']

                        # 检查机型是否匹配
                        is_matching_acft = (acft_cat == 'C' and flight_acft in ['A', 'B', 'C', 'D']) or \
                                           (acft_cat == 'E' and flight_acft in ['E', 'F'])

                        if is_matching_acft and flight_airline == airline and (flight_id, heading) in y:
                            acft_count_expr += y[flight_id, heading]

                    # 未分配航司的航班
                    for i, flight in hour_flights.iterrows():
                        flight_id = flight['ID']
                        flight_airline = result_df.loc[i, '航司']
                        flight_acft = flight['机型']

                        # 检查机型是否匹配
                        is_matching_acft = (acft_cat == 'C' and flight_acft in ['A', 'B', 'C', 'D']) or \
                                           (acft_cat == 'E' and flight_acft in ['E', 'F'])

                        if is_matching_acft and not flight_airline and (flight_id, airline) in z:
                            for heading_item, _ in all_headings:
                                if (flight_id, heading_item, airline) in y and heading_item == heading:
                                    acft_count_expr += y[flight_id, heading_item, airline]

                    # 如果有航班可能被分配到这个组合
                    if acft_count_expr.size() > 0:
                        # 创建变量表示该组合的航班数量
                        key = (airline, heading, acft_cat, hour)
                        var_name = f"future_{airline}_{heading}_{acft_cat}_{hour}"
                        future_distribution[key] = model.addVar(name=var_name, lb=0, vtype=GRB.INTEGER)

                        # 添加约束，确保future_distribution变量等于acft_count_expr
                        model.addConstr(future_distribution[key] == acft_count_expr,
                                        f"future_distr_{airline}_{heading}_{acft_cat}_{hour}")

    # 计算总航班数
    total_flights = day2_departures.shape[0]

    # 计算未来分布与当前分布的偏离度
    for key in set(current_hour_distribution.keys()) | set(future_distribution.keys()):
        current_value = current_hour_distribution.get(key, 0)

        if key in future_distribution:
            # 如果未来分布中有这个组合，计算偏离度
            future_value = future_distribution[key] / total_flights  # 转换为比例

            # 添加偏离度到目标函数（使用绝对偏差）
            deviation_var = model.addVar(name=f"dev_{'_'.join(str(k) for k in key)}", lb=0)

            # 添加约束来定义偏离度变量
            model.addConstr(deviation_var >= future_value - current_value, f"dev_pos_{'_'.join(str(k) for k in key)}")
            model.addConstr(deviation_var >= current_value - future_value, f"dev_neg_{'_'.join(str(k) for k in key)}")

            # 将偏离度添加到目标函数，权重可以根据需要调整
            obj_expr += deviation_var * 100
        else:
            # 如果未来分布中没有这个组合，但当前分布有，则偏离度为当前值
            if current_value > 0:
                obj_expr += current_value * 100

    # 设置目标函数为最小化偏离度
    model.setObjective(obj_expr, GRB.MINIMIZE)

    # 求解模型
    model.optimize()

    # 检查模型是否找到了可行解
    if model.status == GRB.OPTIMAL:
        print(f"{market_type}离港航班分配成功！")

        # 将结果添加到航班数据中
        for i, flight in day2_departures.iterrows():
            flight_id = flight['ID']
            flight_airline = result_df.loc[i, '航司']

            # 对于已分配航司的航班，只需分配航向
            if flight_airline:
                for heading, _ in all_headings:
                    if (flight_id, heading) in y and y[flight_id, heading].X > 0.5:
                        result_df.loc[i, '航向'] = heading
            # 对于未分配航司的航班，需要分配航司和航向
            else:
                for airline, _ in all_airlines:
                    if (flight_id, airline) in z and z[flight_id, airline].X > 0.5:
                        result_df.loc[i, '航司'] = airline

                        for heading, _ in all_headings:
                            if (flight_id, heading, airline) in y and y[flight_id, heading, airline].X > 0.5:
                                result_df.loc[i, '航向'] = heading

        return result_df
    else:
        print(f"{market_type}离港航班分配失败！模型状态：{model.status}")
        return None


def complete_day3_assignments(day2_arr_assignments, departure_flights):
    """
    用日期为2的进港航班信息，补齐日期为3的离港航班信息
    
    Args:
        day2_arr_assignments: 日期为2的进港航班分配结果
        departure_flights: 所有离港航班
    
    Returns:
        补齐后的日期为3的离港航班分配结果
    """
    print("开始补齐日期为3的离港航班分配...")

    # 筛选出日期为3的离港航班
    day3_departures = departure_flights[departure_flights['日期'] == 3].copy()
    day3_departures['航司'] = ""
    day3_departures['航向'] = ""

    # 为每个日期为3的离港航班分配航司和航向
    for i, flight in day3_departures.iterrows():
        flight_id = flight['ID']

        # 查找相同ID的日期为2的进港航班
        matching_arrivals = day2_arr_assignments[day2_arr_assignments['ID'] == flight_id]

        if not matching_arrivals.empty:
            # 从日期为2的进港航班分配结果中获取航司
            day3_departures.loc[i, '航司'] = matching_arrivals.iloc[0]['航司']

        # 为航向随机分配一个与市场类型匹配的值
        market_type = flight['市场']
        if market_type == 'DOM':
            headings = list(HEADINGS_DEP.keys())
        else:
            headings = list(HEADINGS_DEP.keys())

        # 简单地随机分配一个航向
        import random
        day3_departures.loc[i, '航向'] = random.choice(headings)

    return day3_departures


def complete_day1_assignments(day2_arr_assignments, day3_dep_assignments, arrival_flights, departure_flights):
    """
    用（DAY2进港-DAY3离港）航班信息，补齐（DAY1进港-DAY2离港航班信息）
    
    Args:
        day2_arr_assignments: 日期为2的进港航班分配结果
        day3_dep_assignments: 日期为3的离港航班分配结果
        arrival_flights: 所有进港航班
        departure_flights: 所有离港航班
    
    Returns:
        补齐后的日期为1的进港航班分配结果
    """
    print("开始补齐日期为1的进港航班分配...")

    # 筛选出日期为1的进港航班
    day1_arrivals = arrival_flights[arrival_flights['日期'] == 1].copy()
    day1_arrivals['航司'] = ""
    day1_arrivals['航向'] = ""

    # 筛选出日期为2的离港航班
    day2_departures = departure_flights[departure_flights['日期'] == 2]

    # 为每个日期为1的进港航班分配航司和航向
    for i, flight in day1_arrivals.iterrows():
        flight_id = flight['ID']

        # 查找相同ID的日期为2的离港航班
        matching_departures = day2_departures[day2_departures['ID'] == flight_id]

        if not matching_departures.empty:
            dep_id = matching_departures.iloc[0].name

            # 从日期为2的离港航班分配结果中获取航司
            if dep_id in day3_dep_assignments.index:
                day1_arrivals.loc[i, '航司'] = day3_dep_assignments.loc[dep_id, '航司']

            # 为航向随机分配一个与市场类型匹配的值
            market_type = flight['市场']
            if market_type == 'DOM':
                headings = list(HEADINGS_ARR.keys())
            else:
                headings = list(HEADINGS_ARR.keys())

            # 简单地随机分配一个航向
            import random
            day1_arrivals.loc[i, '航向'] = random.choice(headings)

    return day1_arrivals


def main():
    # 加载数据
    current_status, arrival_flights, departure_flights, hourly_dom_stats, hourly_int_stats, hourly_dom_dep_stats, hourly_int_dep_stats = load_data()

    # 分配国内进港航班，传入离港航班数据
    dom_arr_assignments, dom_dep_counts, int_dep_counts = assign_arrival_flights(
        current_status, arrival_flights, 'DOM', hourly_dom_stats, hourly_int_stats, departure_flights
    )

    # 分配国际进港航班，传入离港航班数据和之前的分配结果
    int_arr_assignments, _, _ = assign_arrival_flights(
        current_status, arrival_flights, 'INT', hourly_dom_stats, hourly_int_stats,
        departure_flights, dom_dep_counts, int_dep_counts
    )

    # 合并进港航班分配结果
    all_day2_arr_assignments = pd.concat([dom_arr_assignments, int_arr_assignments])

    # 分配国内离港航班
    dom_dep_assignments = assign_departure_flights(all_day2_arr_assignments, departure_flights, current_status, 'DOM',
                                                   hourly_dom_dep_stats, hourly_int_dep_stats)

    # 分配国际离港航班
    int_dep_assignments = assign_departure_flights(all_day2_arr_assignments, departure_flights, current_status, 'INT',
                                                   hourly_dom_dep_stats, hourly_int_dep_stats)

    # 合并离港航班分配结果
    all_day2_dep_assignments = pd.concat([dom_dep_assignments, int_dep_assignments])

    # 补齐日期为1的进港航班分配
    day1_arr_assignments = complete_day1_assignments(all_day2_arr_assignments, all_day2_dep_assignments,
                                                     arrival_flights, departure_flights)

    # 补齐日期为3的离港航班分配
    day3_dep_assignments = complete_day3_assignments(all_day2_arr_assignments, departure_flights)

    # 合并所有分配结果
    all_arr_assignments = pd.concat([day1_arr_assignments, all_day2_arr_assignments])
    all_dep_assignments = pd.concat([all_day2_dep_assignments, day3_dep_assignments])

    # 将结果写回Excel文件
    with pd.ExcelWriter(OUTPUT_FILE) as writer:
        all_arr_assignments.to_excel(writer, sheet_name="到达航班", index=False)
        all_dep_assignments.to_excel(writer, sheet_name="出发航班", index=False)

    print(f"分配完成！结果已保存到 {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
