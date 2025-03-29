import pandas as pd
import numpy as np
import gurobipy as gp
from gurobipy import GRB
import os
from dataset import DOM_AIRLINES, INT_AIRLINES, HEADINGS_ARR, HEADINGS_DEP

# 文件路径
CURRENT_STATUS_FILE = "2024_现状.xlsx"
PAIRING_FILE = "final_pairing_processed.xlsx"
OUTPUT_FILE = "final_pairing_processed_with_assignments.xlsx"

def load_data():
    """加载所有需要的数据"""
    # 加载现状数据
    current_status = pd.read_excel(CURRENT_STATUS_FILE)
    
    # 加载配对数据
    arrival_flights = pd.read_excel(PAIRING_FILE, sheet_name="到达航班")
    departure_flights = pd.read_excel(PAIRING_FILE, sheet_name="出发航班")
    
    return current_status, arrival_flights, departure_flights

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

def assign_arrival_flights(current_status, arrival_flights, market_type):
    """
    为进港航班分配航司和航向
    
    Args:
        current_status: 现状数据
        arrival_flights: 需要分配的进港航班
        market_type: 'DOM' 或 'INT'
    
    Returns:
        带有航司和航向分配的进港航班数据
    """
    print(f"开始分配{market_type}进港航班...")

    main_headings = get_main_headings(current_status)

    # 筛选出日期为2的进港航班
    day2_arrivals = arrival_flights[(arrival_flights['日期'] == 2) & (arrival_flights['市场'] == market_type)]

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
    current_hour_airline_heading_count = current_arrivals.groupby(['小时', 'AirlineGroup', 'Routing']).size().reset_index(name='count')
    current_hour_airline_heading_acft_count = current_arrivals.groupby(['小时', 'AirlineGroup', 'Routing', 'Acft Cat']).size().reset_index(name='count')
    
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
                if market_type == 'INT' and heading_type == '绝对远程' and base_type == '非主基地' and airline not in ['东航集团', '南航集团', '海航集团']:
                    continue
                
                # 2. 绝对远程航向分配宽体机
                if market_type == 'INT' and heading_type == '绝对远程' and flight_acft not in ['E', 'F']:
                    continue
                
                x[flight_id, airline, heading] = model.addVar(vtype=GRB.BINARY, name=f"x_{flight_id}_{airline}_{heading}")
    
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


    # # 目标函数：最大化与现状分布的相似性
    # # 这里我们使用一个简单的方法，尝试让每个小时内各航司和航向的分布与现状相似
    # obj_expr = gp.LinExpr()
    #
    # # 添加航司×时间的相似性
    # for _, row in current_hour_airline_count.iterrows():
    #     hour = row['小时']
    #     airline = row['AirlineGroup']
    #     count = row['count']
    #
    #     # 计算在新数据中该小时该航司的航班数量
    #     hour_airline_flights = day2_arrivals[day2_arrivals['小时'] == hour]
    #
    #     if not hour_airline_flights.empty and (airline, _) in all_airlines:
    #         for i, flight in hour_airline_flights.iterrows():
    #             flight_id = flight['ID']
    #             for heading, _ in all_headings:
    #                 if (flight_id, airline, heading) in x:
    #                     obj_expr += x[flight_id, airline, heading] * count
    #
    # # 修改航司×时间×航向的相似性部分，使用主航向信息
    # for _, row in current_hour_airline_heading_count.iterrows():
    #     hour = row['小时']
    #     airline = row['AirlineGroup']
    #     heading = row['Routing']
    #     count = row['count']
    #
    #     # 计算在新数据中该小时该航司该航向的航班数量
    #     hour_flights = day2_arrivals[day2_arrivals['小时'] == hour]
    #
    #     if not hour_flights.empty and (airline, _) in all_airlines and (heading, _) in all_headings:
    #         for i, flight in hour_flights.iterrows():
    #             flight_id = flight['ID']
    #             if (flight_id, airline, heading) in x:
    #                 # 主航向给予更高的权重
    #                 key = (airline, market_type, 'Arrival')
    #                 is_main_heading = key in main_headings and main_headings[key][0] == heading
    #                 weight = 3 if is_main_heading else 1
    #                 obj_expr += x[flight_id, airline, heading] * count * weight
    #
    #
    # # 添加航司×时间×航向×机型的相似性
    # for _, row in current_hour_airline_heading_acft_count.iterrows():
    #     hour = row['小时']
    #     airline = row['AirlineGroup']
    #     heading = row['Routing']
    #     acft = row['Acft Cat']
    #     count = row['count']
    #
    #     # 计算在新数据中该小时该航司该航向该机型的航班数量
    #     hour_acft_flights = day2_arrivals[(day2_arrivals['小时'] == hour) & (day2_arrivals['机型'] == acft)]
    #
    #     if not hour_acft_flights.empty and (airline, _) in all_airlines and (heading, _) in all_headings:
    #         for i, flight in hour_acft_flights.iterrows():
    #             flight_id = flight['ID']
    #             if (flight_id, airline, heading) in x:
    #                 # 主基地航司的宽体机比例增幅较大
    #                 weight = 1
    #                 if acft in ['E', 'F'] and airline in ['川航集团', '国航集团']:
    #                     weight = 3
    #                 obj_expr += x[flight_id, airline, heading] * count * weight
    
    model.setObjective(0, GRB.MAXIMIZE)
    
    # 求解模型
    model.optimize()
    
    # 检查模型是否找到了可行解
    if model.status == GRB.OPTIMAL:
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
        
        return result_df
    else:
        print(f"{market_type}进港航班分配失败！模型状态：{model.status}")
        return None

def assign_departure_flights(arrival_assignments, departure_flights, current_status, market_type):
    """
    为离港航班分配航司和航向
    
    Args:
        arrival_assignments: 已分配的进港航班
        departure_flights: 需要分配的离港航班
        current_status: 现状数据
        market_type: 'DOM' 或 'INT'
    
    Returns:
        带有航司和航向分配的离港航班数据
    """
    print(f"开始分配{market_type}离港航班...")
    main_headings = get_main_headings(current_status)

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
    current_hour_airline_heading_count = current_departures.groupby(['小时', 'AirlineGroup', 'Routing']).size().reset_index(name='count')
    current_hour_airline_heading_acft_count = current_departures.groupby(['小时', 'AirlineGroup', 'Routing', 'Acft Cat']).size().reset_index(name='count')

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
            # 获取进港航班的日期
            arrival_date = matching_arrivals.iloc[0].get('日期')

            # 如果进港航班日期为2，则遵循"离港航司=进港航司"的规则
            # 如果进港航班日期为1，则不需要遵循该规则，航司将在后续步骤中分配
            if arrival_date == 2:
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
            airline_base_type = next((base_type for airline, base_type in all_airlines if airline == flight_airline), None)

            for heading, heading_type in all_headings:
                # 检查是否满足约束条件
                # 1. 绝对远程航向只分配给主基地航司或东航南航海航集团
                if market_type == 'INT' and heading_type == '绝对远程' and airline_base_type == '非主基地' and flight_airline not in ['东航集团', '南航集团', '海航集团']:
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
                    if market_type == 'INT' and heading_type == '绝对远程' and base_type == '非主基地' and airline not in ['东航集团', '南航集团', '海航集团']:
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
                                                         if (flight_id := day2_departures.loc[i, 'ID'], main_heading) in y)

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
                                                           if (flight_id := day2_departures.loc[i, 'ID'], heading) in y)

                        unassigned_heading_expr = gp.quicksum(y[flight_id, heading, airline]
                                                             for i, flight in day2_departures.iterrows()
                                                             if not result_df.loc[i, '航司']
                                                             and (flight_id := flight['ID'], heading, airline) in y)

                        model.addConstr(
                            (assigned_main_heading_expr + unassigned_main_heading_expr) >=
                            (assigned_heading_expr + unassigned_heading_expr),
                            f"main_heading_dominance_{airline}_{market_type}_{heading}_Departure"
                        )
    # # 目标函数：最大化与现状分布的相似性
    # obj_expr = gp.LinExpr()
    #
    # # 修改航司×时间×航向的相似性部分，使用主航向信息
    # for _, row in current_hour_airline_heading_count.iterrows():
    #     hour = row['小时']
    #     airline = row['AirlineGroup']
    #     heading = row['Routing']
    #     count = row['count']
    #
    #     # 计算在新数据中该小时该航司该航向的航班数量
    #     hour_airline_flights = day2_departures[day2_departures['小时'] == hour]
    #
    #     for i, flight in hour_airline_flights.iterrows():
    #         flight_id = flight['ID']
    #         flight_airline = result_df.loc[i, '航司']
    #
    #         # 主航向给予更高的权重
    #         key = (airline, market_type, 'Departure')
    #         is_main_heading = key in main_headings and main_headings[key][0] == heading
    #         weight = 3 if is_main_heading else 1
    #
    #         if flight_airline == airline and (flight_id, heading) in y:
    #             obj_expr += y[flight_id, heading] * count * weight
    #         elif not flight_airline and (flight_id, airline) in z:
    #             for heading_item, _ in all_headings:
    #                 if (flight_id, heading_item, airline) in y and heading_item == heading:
    #                     obj_expr += y[flight_id, heading_item, airline] * count * weight
    #
    # # 添加航司×时间×航向×机型的相似性
    # for _, row in current_hour_airline_heading_acft_count.iterrows():
    #     hour = row['小时']
    #     airline = row['AirlineGroup']
    #     heading = row['Routing']
    #     acft = row['Acft Cat']
    #     count = row['count']
    #
    #     # 计算在新数据中该小时该航司该航向该机型的航班数量
    #     hour_acft_airline_flights = day2_departures[(day2_departures['小时'] == hour) & (day2_departures['机型'] == acft)]
    #
    #     for i, flight in hour_acft_airline_flights.iterrows():
    #         flight_id = flight['ID']
    #         flight_airline = result_df.loc[i, '航司']
    #
    #         if flight_airline == airline and (flight_id, heading) in y:
    #             # 主基地航司的宽体机比例增幅较大
    #             weight = 1
    #             if acft in ['E', 'F'] and airline in ['川航集团', '国航集团']:
    #                 weight = 3
    #             obj_expr += y[flight_id, heading] * count * weight
    #         elif not flight_airline and (flight_id, airline) in z:
    #             for heading_item, _ in all_headings:
    #                 if (flight_id, heading_item, airline) in y and heading_item == heading:
    #                     # 主基地航司的宽体机比例增幅较大
    #                     weight = 1
    #                     if acft in ['E', 'F'] and airline in ['川航集团', '国航集团']:
    #                         weight = 3
    #                     obj_expr += y[flight_id, heading_item, airline] * count * weight

    model.setObjective(0, GRB.MAXIMIZE)

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
    current_status, arrival_flights, departure_flights = load_data()
    
    # 分配国内进港航班
    dom_arr_assignments = assign_arrival_flights(current_status, arrival_flights, 'DOM')
    
    # 分配国际进港航班
    int_arr_assignments = assign_arrival_flights(current_status, arrival_flights, 'INT')
    
    # 合并进港航班分配结果
    all_day2_arr_assignments = pd.concat([dom_arr_assignments, int_arr_assignments])
    
    # 分配国内离港航班
    dom_dep_assignments = assign_departure_flights(dom_arr_assignments, departure_flights, current_status, 'DOM')
    
    # 分配国际离港航班
    int_dep_assignments = assign_departure_flights(int_arr_assignments, departure_flights, current_status, 'INT')
    
    # 合并离港航班分配结果
    all_day2_dep_assignments = pd.concat([dom_dep_assignments, int_dep_assignments])
    
    # 补齐日期为1的进港航班分配
    day1_arr_assignments = complete_day1_assignments(all_day2_arr_assignments, all_day2_dep_assignments, arrival_flights, departure_flights)
    
    # 合并所有分配结果
    all_arr_assignments = pd.concat([day1_arr_assignments, all_day2_arr_assignments])
    
    # 将结果写回Excel文件
    with pd.ExcelWriter(OUTPUT_FILE) as writer:
        all_arr_assignments.to_excel(writer, sheet_name="到达航班", index=False)
        all_day2_dep_assignments.to_excel(writer, sheet_name="出发航班", index=False)
    
    print(f"分配完成！结果已保存到 {OUTPUT_FILE}")

if __name__ == "__main__":
    main()