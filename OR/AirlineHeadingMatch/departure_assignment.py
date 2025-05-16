import gurobipy as gp
import pandas as pd
from gurobipy import GRB

from OR.AirlineHeadingMatch.utils import get_main_headings, get_hour_from_time, diagnose_infeasibility
from config import DOM_DEP_GAP, INT_DEP_GAP, DOM_DEP_WAVE_BIAS, INT_DEP_WAVE_BIAS, DOM_DEP_WIDE_BIAS, INT_DEP_WIDE_BIAS
from excel_to_dataset import DOM_AIRLINES, INT_AIRLINES, HEADINGS_DEP, AIRLINES_WIDE, ABSOLUTE_LONG_ROUTING, \
    MAIN_HEADING_EXCEPTION_AIRLINES, WAVE_EXCEPTION_AIRLINES, DEP_DOM_WIDE_EXCEPTION_ROUTING, \
    DEP_INT_WIDE_EXCEPTION_ROUTING, DEP_INT_WIDE_UP_ROUTING, DEP_DOM_WIDE_UP_ROUTING


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
    day2_departures = departure_flights[
        (departure_flights['日期'] == 2) &
        (departure_flights['市场'] == market_type)
        ].copy()

    # 筛选现状数据中的离港航班
    current_departures = current_status[
        (current_status['Direction'] == 'Departure') &
        (current_status['flightnature'] == market_type)
        ].copy()

    # 为每个航班添加小时列
    day2_departures.loc[:, '小时'] = day2_departures['时间'].apply(get_hour_from_time)
    current_departures.loc[:, '小时'] = (current_departures['Minute for rolling'] // 60) % 24

    # 获取现状中各航司的航向数据
    airline_heading_pairs = current_departures.groupby(['AirlineGroup', 'Routing']).size().reset_index()
    valid_airline_heading_pairs = set(zip(airline_heading_pairs['AirlineGroup'], airline_heading_pairs['Routing']))
    
    # 获取所有可能的航司和航向
    if market_type == 'DOM':
        airlines_data = DOM_AIRLINES
        headings_data = HEADINGS_DEP
        wide_up_routings = DEP_DOM_WIDE_UP_ROUTING
    else:
        airlines_data = INT_AIRLINES
        headings_data = HEADINGS_DEP
        wide_up_routings = DEP_INT_WIDE_UP_ROUTING

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
        is_wide_body = flight_acft in ['E', 'F']

        # 如果已经分配了航司
        if flight_airline:
            airline_base_type = next((base_type for airline, base_type in all_airlines if airline == flight_airline),
                                     None)

            for heading, heading_type in all_headings:
                # 检查是否满足约束条件
                # 1. 绝对远程航向只分配给主基地航司或东航南航海航集团
                if market_type == 'INT' and heading_type == '绝对远程' and airline_base_type == '非主基地' and flight_airline not in ABSOLUTE_LONG_ROUTING:
                    continue

                # 2. 绝对远程航向分配宽体机
                if market_type == 'INT' and heading_type == '绝对远程' and flight_acft not in ['E', 'F']:
                    continue

                # 3. 特定航向不能有宽体机
                if is_wide_body:
                    if market_type == 'DOM' and heading in DEP_DOM_WIDE_EXCEPTION_ROUTING:
                        continue
                    if market_type == 'INT' and heading in DEP_INT_WIDE_EXCEPTION_ROUTING:
                        continue

                # 4. 新增约束：只有当航司在现状中存在该航向时，才能分配
                # 特例：国际离港的SAGPI航向可以分配给所有航司
                if flight_airline != '其它' and (flight_airline, heading) not in valid_airline_heading_pairs:
                    if not (market_type == 'INT' and heading == 'SAGPI'):
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
                    if market_type == 'INT' and heading_type == '绝对远程' and base_type == '非主基地' and airline not in ABSOLUTE_LONG_ROUTING:
                        continue

                    # 2. 绝对远程航向分配宽体机
                    if market_type == 'INT' and heading_type == '绝对远程' and flight_acft not in ['E', 'F']:
                        continue
                    # 3. 特定航向不能有宽体机
                    if is_wide_body:
                        if market_type == 'DOM' and heading in DEP_DOM_WIDE_EXCEPTION_ROUTING:
                            continue
                        if market_type == 'INT' and heading in DEP_INT_WIDE_EXCEPTION_ROUTING:
                            continue

                    # 4. 新增约束：只有当航司在现状中存在该航向时，才能分配
                    # 特例：国际离港的SAGPI航向可以分配给所有航司
                    if airline != '其它' and (airline, heading) not in valid_airline_heading_pairs:
                        if not (market_type == 'INT' and heading == 'SAGPI'):
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
            bias = 0  # 使用DOM的偏差值
        else:
            quota = HEADINGS_DEP[heading]['INT']
            bias = 0  # 使用INT的偏差值

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

        model.addConstr(assigned_expr + unassigned_expr >= quota - bias, f"heading_quota_min_{heading}")
        model.addConstr(assigned_expr + unassigned_expr <= quota + bias, f"heading_quota_max_{heading}")

    # 主航向比例约束：确保各航司的主航向比例不低于现状
    for airline, _ in all_airlines:
        if airline in MAIN_HEADING_EXCEPTION_AIRLINES:
            continue

        # if market_type == 'INT' and airline =='海航集团':
        #     continue

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

                # 确保主航向仍然是航班量最多的航向
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

    # 新增约束：确保各航司集团的未来小时波形大于现状波形
    # 按航司集团和小时统计现状航班数量
    airline_hour_counts = current_departures.groupby(['AirlineGroup', '小时']).size().reset_index(name='count')

    # 为每个航司和每个小时添加约束
    for airline, _ in all_airlines:
        if airline in WAVE_EXCEPTION_AIRLINES:
            continue

        for hour in range(24):
            # 获取现状中该航司该小时的航班数量
            current_count = airline_hour_counts[
                (airline_hour_counts['AirlineGroup'] == airline) &
                (airline_hour_counts['小时'] == hour)
                ]['count'].sum() if not airline_hour_counts[
                (airline_hour_counts['AirlineGroup'] == airline) &
                (airline_hour_counts['小时'] == hour)
                ].empty else 0

            # 计算未来分布中该航司该小时的航班数量 - 已分配航司的航班
            assigned_count_expr = gp.quicksum(
                y[flight_id, heading]
                for i, flight in day2_departures.iterrows()
                if flight['小时'] == hour and result_df.loc[i, '航司'] == airline
                for heading, _ in all_headings
                if (flight_id := flight['ID'], heading) in y
            )

            # 计算未来分布中该航司该小时的航班数量 - 未分配航司的航班
            unassigned_count_expr = gp.quicksum(
                y[flight_id, heading, airline]
                for i, flight in day2_departures.iterrows()
                if flight['小时'] == hour and not result_df.loc[i, '航司']
                for heading, _ in all_headings
                if (flight_id := flight['ID'], heading, airline) in y
            )

            # 添加约束：未来分布不低于现状
            if current_count > 0:
                # 波形偏移
                if market_type == 'DOM':
                    bias = DOM_DEP_WAVE_BIAS
                else:
                    bias = INT_DEP_WAVE_BIAS
                model.addConstr(
                    assigned_count_expr + unassigned_count_expr >= current_count - bias,
                    f"hourly_wave_{airline}_{hour}_{market_type}_Departure"
                )

    # 添加宽体机配额约束
    for airline, _ in all_airlines:
        # 获取该航司在该市场类型的宽体机配额
        wide_body_quota = AIRLINES_WIDE.get(airline, {}).get(market_type, 0)

        # 计算已分配给该航司的宽体机航班数量
        assigned_wide_body_count = sum(
            1 for i, flight in day2_departures.iterrows()
            if result_df.loc[i, '航司'] == airline and flight['机型'] in ['E', 'F']
        )

        # 计算未分配航司的宽体机航班中，分配给该航司的数量表达式
        unassigned_wide_body_expr = gp.quicksum(
            z[flight_id, airline]
            for i, flight in day2_departures.iterrows()
            if not result_df.loc[i, '航司'] and flight['机型'] in ['E', 'F']
            and (flight_id := flight['ID'], airline) in z
        )

        # 添加约束：宽体机总数量必须等于配额
        # 宽体偏移
        if market_type == 'DOM':
            bias = DOM_DEP_WIDE_BIAS
            if airline == '其它':
                bias = 6
        else:
            bias = INT_DEP_WIDE_BIAS
        model.addConstr(
            assigned_wide_body_count + unassigned_wide_body_expr >= wide_body_quota - bias,
            f"wide_body_quota_{airline}_{market_type}_min_Departure"
        )
        model.addConstr(
            assigned_wide_body_count + unassigned_wide_body_expr <= wide_body_quota + bias,
            f"wide_body_quota_{airline}_{market_type}_max_Departure"
        )

    # 添加宽体机比例升高约束
    if wide_up_routings:  # 如果有需要提高宽体机比例的航向
        # 计算现状中各航向的宽体机比例
        current_wide_ratio = {}
        for heading in wide_up_routings:
            # 筛选该航向的航班
            heading_flights = current_departures[current_departures['Routing'] == heading]
            if not heading_flights.empty:
                # 计算宽体机航班数量
                wide_count = heading_flights[heading_flights['Acft Cat'].isin(['E', 'F'])].shape[0]
                total_count = heading_flights.shape[0]
                if total_count > 0:
                    current_wide_ratio[heading] = wide_count / total_count
                else:
                    current_wide_ratio[heading] = 0
            else:
                current_wide_ratio[heading] = 0

        # 为每个需要提高宽体机比例的航向添加约束
        for heading in wide_up_routings:
            if heading in current_wide_ratio:
                # 计算未来该航向的总航班数 - 已分配航司的航班
                assigned_total_expr = gp.quicksum(
                    y[flight_id, heading]
                    for i, flight in day2_departures.iterrows()
                    if result_df.loc[i, '航司'] and (flight_id := flight['ID'], heading) in y
                )

                # 计算未来该航向的宽体机航班数 - 已分配航司的航班
                assigned_wide_expr = gp.quicksum(
                    y[flight_id, heading]
                    for i, flight in day2_departures.iterrows()
                    if result_df.loc[i, '航司'] and (flight_id := flight['ID'], heading) in y
                    and flight['机型'] in ['E', 'F']
                )

                # 计算未来该航向的总航班数 - 未分配航司的航班
                unassigned_total_expr = gp.quicksum(
                    y[flight_id, heading, airline]
                    for i, flight in day2_departures.iterrows()
                    if not result_df.loc[i, '航司']
                    for airline, _ in all_airlines
                    if (flight_id := flight['ID'], heading, airline) in y
                )

                # 计算未来该航向的宽体机航班数 - 未分配航司的航班
                unassigned_wide_expr = gp.quicksum(
                    y[flight_id, heading, airline]
                    for i, flight in day2_departures.iterrows()
                    if not result_df.loc[i, '航司']
                    for airline, _ in all_airlines
                    if (flight_id := flight['ID'], heading, airline) in y
                    and flight['机型'] in ['E', 'F']
                )

                # 添加约束：未来宽体机比例 > 现状宽体机比例
                total_expr = assigned_total_expr + unassigned_total_expr
                wide_expr = assigned_wide_expr + unassigned_wide_expr

                if total_expr.size() > 0:  # 确保有航班分配给该航向
                    model.addConstr(
                        wide_expr - 0.0001 >= current_wide_ratio[heading] * total_expr,
                        f"wide_body_ratio_increase_{heading}_{market_type}_Departure"
                    )

    # 添加基于小时统计数据的约束和目标函数
    # 目标函数：最小化未来分布与当前分布的偏离度

    # ========== 新增：多目标优化 ==========
    # 第一目标：最小化各航司集团（除其它）未来波形与现状波形的差异
    airline_wave_deviation = gp.LinExpr()
    for airline, _ in all_airlines:
        if airline == '其它':
            continue
        for hour in range(24):
            # 现状
            current_count = airline_hour_counts[
                (airline_hour_counts['AirlineGroup'] == airline) &
                (airline_hour_counts['小时'] == hour)
                ]['count'].sum() if not airline_hour_counts[
                (airline_hour_counts['AirlineGroup'] == airline) &
                (airline_hour_counts['小时'] == hour)
                ].empty else 0

            # 未来
            assigned_count_expr = gp.quicksum(
                y[flight_id, heading]
                for i, flight in day2_departures.iterrows()
                if flight['小时'] == hour and result_df.loc[i, '航司'] == airline
                for heading, _ in all_headings
                if (flight_id := flight['ID'], heading) in y
            )
            unassigned_count_expr = gp.quicksum(
                y[flight_id, heading, airline]
                for i, flight in day2_departures.iterrows()
                if flight['小时'] == hour and not result_df.loc[i, '航司']
                for heading, _ in all_headings
                if (flight_id := flight['ID'], heading, airline) in y
            )
            future_count_expr = assigned_count_expr + unassigned_count_expr

            # 绝对偏差
            dev_var = model.addVar(lb=0, name=f"wave_dev_{airline}_{hour}")
            model.addConstr(dev_var >= future_count_expr - current_count, f"wave_dev_pos_{airline}_{hour}")
            model.addConstr(dev_var >= current_count - future_count_expr, f"wave_dev_neg_{airline}_{hour}")
            dev_var = dev_var ** 2
            airline_wave_deviation += dev_var

    # 第二目标：原有目标函数
    obj_expr = gp.LinExpr()

    # 首先，计算当前分布的小时分布情况
    current_hour_distribution = {}
    for _, row in hourly_stats.iterrows():
        airline = row['AirlineGroup']
        routing = row['Routing']
        acft_cat = row['Acft Cat']

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

    # ========== 设置多目标 ==========

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
    model.setObjective(airline_wave_deviation * 10000 + obj_expr, GRB.MINIMIZE)
    if market_type == 'DOM':
        model.Params.MIPGap = DOM_DEP_GAP  # 设置Gap（相对间隙）
    else:
        model.Params.MIPGap = INT_DEP_GAP

    # 求解模型
    model.optimize()

    # 检查模型是否找到了可行解
    if model.status == GRB.OPTIMAL:
        print("-----------------------------------------------------------------------")
        print(f"{market_type}离港航班分配成功！")
        print("-----------------------------------------------------------------------")

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
        diagnose_infeasibility(model)
        print("-----------------------------------------------------------------------")
        print(f"{market_type}离港航班分配失败！模型不可行。")
        print("-----------------------------------------------------------------------")
        return None
