import gurobipy as gp
import pandas as pd
from gurobipy import GRB

from OR.AirlineHeadingMatch.utils import get_main_headings, get_hour_from_time, diagnose_infeasibility
from config import DOM_ARR_GAP, INT_ARR_GAP, DOM_ARR_WAVE_BIAS, INT_ARR_WAVE_BIAS, DOM_ARR_WIDE_BIAS, INT_ARR_WIDE_BIAS
from excel_to_dataset import DOM_AIRLINES, INT_AIRLINES, HEADINGS_ARR, AIRLINES_WIDE, ABSOLUTE_LONG_ROUTING, \
    MAIN_HEADING_EXCEPTION_AIRLINES, WAVE_EXCEPTION_AIRLINES, ARR_DOM_WIDE_EXCEPTION_ROUTING, \
    ARR_INT_WIDE_EXCEPTION_ROUTING,ARR_INT_WIDE_UP_ROUTING,ARR_DOM_WIDE_UP_ROUTING


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

    day2_arrivals = arrival_flights[(arrival_flights['日期'] == 2) & (arrival_flights['市场'] == market_type)].copy()

    # 筛选现状数据中的进港航班
    current_arrivals = current_status[
        (current_status['Direction'] == 'Arrival') &
        (current_status['flightnature'] == market_type)
        ].copy()

    # 为每个航班添加小时列
    day2_arrivals.loc[:, '小时'] = day2_arrivals['时间'].apply(get_hour_from_time)
    current_arrivals.loc[:, '小时'] = (current_arrivals['Minute for rolling'] // 60) % 24

    # 获取现状中各航司的航向数据
    airline_heading_pairs = current_arrivals.groupby(['AirlineGroup', 'Routing']).size().reset_index()
    valid_airline_heading_pairs = set(zip(airline_heading_pairs['AirlineGroup'], airline_heading_pairs['Routing']))
    
    # 获取所有可能的航司和航向
    if market_type == 'DOM':
        airlines_data = DOM_AIRLINES
        headings_data = HEADINGS_ARR
        wide_up_routings = ARR_DOM_WIDE_UP_ROUTING
    else:
        airlines_data = INT_AIRLINES
        headings_data = HEADINGS_ARR
        wide_up_routings = ARR_INT_WIDE_UP_ROUTING

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
        is_wide_body = flight_acft in ['E','F']

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
                    if market_type == 'DOM' and heading in ARR_DOM_WIDE_EXCEPTION_ROUTING:
                        continue
                    if market_type == 'INT' and heading in ARR_INT_WIDE_EXCEPTION_ROUTING:
                        continue
                
                # 4. 新增约束：只有当航司在现状中存在该航向时，才能分配
                if airline != '其它' and (airline, heading) not in valid_airline_heading_pairs:
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
        if airline in MAIN_HEADING_EXCEPTION_AIRLINES:
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

    # 每个航司分配的进港航班对应的离港航班总量不超过配额
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

    # 新增约束：确保各航司集团的未来小时波形大于现状波形
    # 按航司集团和小时统计现状航班数量
    airline_hour_counts = current_arrivals.groupby(['AirlineGroup', '小时']).size().reset_index(name='count')

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

            # 计算未来分布中该航司该小时的航班数量
            future_count_expr = gp.quicksum(
                x[flight_id, airline, heading]
                for i, flight in day2_arrivals.iterrows()
                if flight['小时'] == hour
                for heading, _ in all_headings
                if (flight_id := flight['ID'], airline, heading) in x
            )

            # 添加约束：未来分布不低于现状
            if current_count > 0:
                if market_type == 'DOM':
                    bias = DOM_ARR_WAVE_BIAS
                else:
                    bias = INT_ARR_WAVE_BIAS
                model.addConstr(
                    future_count_expr >= current_count - bias,
                    f"hourly_wave_{airline}_{hour}_{market_type}_Arrival"
                )

    # 添加宽体机配额约束
    for airline, _ in all_airlines:
        # 获取该航司在该市场类型的宽体机配额
        wide_body_quota = AIRLINES_WIDE.get(airline, {}).get(market_type, 0)

        # 计算分配给该航司的宽体机航班数量
        wide_body_expr = gp.quicksum(
            x[flight_id, airline, heading]
            for i, flight in day2_arrivals.iterrows()
            for heading, _ in all_headings
            if (flight_id := flight['ID'], airline, heading) in x
            and flight['机型'] in ['E', 'F']  # E和F代表宽体机
        )

        # 添加约束：宽体机数量必须等于配额
        if market_type == 'DOM':
            bias = DOM_ARR_WIDE_BIAS
        else:
            bias = INT_ARR_WIDE_BIAS
        model.addConstr(
            wide_body_expr >= wide_body_quota - bias ,
            f"wide_body_quota_{airline}_{market_type}_min_Arrival"
        )
        model.addConstr(
            wide_body_expr <= wide_body_quota + bias ,
            f"wide_body_quota_{airline}_{market_type}_max_Arrival"
        )

    # 添加宽体机比例升高约束
    if wide_up_routings:  # 如果有需要提高宽体机比例的航向
        # 计算现状中各航向的宽体机比例
        current_wide_ratio = {}
        for heading in wide_up_routings:
            # 筛选该航向的航班
            heading_flights = current_arrivals[current_arrivals['Routing'] == heading]
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
                # 计算未来该航向的总航班数
                future_total_expr = gp.quicksum(
                    x[flight_id, airline, heading]
                    for i, flight in day2_arrivals.iterrows()
                    for airline, _ in all_airlines
                    if (flight_id := flight['ID'], airline, heading) in x
                )

                # 计算未来该航向的宽体机航班数
                future_wide_expr = gp.quicksum(
                    x[flight_id, airline, heading]
                    for i, flight in day2_arrivals.iterrows()
                    for airline, _ in all_airlines
                    if (flight_id := flight['ID'], airline, heading) in x
                    and flight['机型'] in ['E', 'F']
                )

                # 添加约束：未来宽体机比例 > 现状宽体机比例
                if future_total_expr.size() > 0:  # 确保有航班分配给该航向
                    model.addConstr(
                        future_wide_expr >= current_wide_ratio[heading] * future_total_expr,
                        f"wide_body_ratio_increase_{heading}_{market_type}_Arrival"
                    )

    # 添加基于小时统计数据的约束和目标函数
    # 目标函数：最大化与小时统计数据的匹配度

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
            future_count_expr = gp.quicksum(
                x[flight_id, airline, heading]
                for i, flight in day2_arrivals.iterrows()
                if flight['小时'] == hour
                for heading, _ in all_headings
                if (flight_id := flight['ID'], airline, heading) in x
            )

            # 绝对偏差
            dev_var = model.addVar(lb=0, name=f"wave_dev_{airline}_{hour}")
            model.addConstr(dev_var >= future_count_expr - current_count, f"wave_dev_pos_{airline}_{hour}")
            model.addConstr(dev_var >= current_count - future_count_expr, f"wave_dev_neg_{airline}_{hour}")
            dev_var = dev_var ** 2
            airline_wave_deviation += dev_var

    # 第二目标：原有目标函数
    obj_expr = gp.LinExpr()

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
    # ========== 设置多目标 ==========
    model.setObjective(airline_wave_deviation * 10000 + obj_expr, GRB.MINIMIZE)
    if market_type == 'DOM':
        model.Params.MIPGap = DOM_ARR_GAP  # 设置Gap（相对间隙）
    else:
        model.Params.MIPGap = INT_ARR_GAP
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
        print("-----------------------------------------------------------------------")
        print(f"{market_type}进港航班分配成功！")
        print("-----------------------------------------------------------------------")

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
        diagnose_infeasibility(model)
        print("-----------------------------------------------------------------------")
        print(f"{market_type}进港航班分配失败！模型不可行。")
        print("-----------------------------------------------------------------------")
        return None, {}, {}
