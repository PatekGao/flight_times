import pandas as pd

from OR.AirlineHeadingMatch.dataset import HEADINGS_DEP, HEADINGS_ARR

# 文件路径
CURRENT_STATUS_FILE = "2024_现状.xlsx"
PAIRING_FILE = "final_pairing_processed.xlsx"
ARR_HOURLY_STATS_FILE = "到达航班比例表.xlsx"  # 新增文件路径
DEP_HOURLY_STATS_FILE = "出发航班比例表.xlsx"

def diagnose_infeasibility(model):
    """
    诊断模型不可行的原因

    Args:
        model: Gurobi模型

    Returns:
        None，但会打印出不可行约束的信息
    """
    print("模型不可行，开始诊断...")

    # 计算不可行子系统 (IIS)
    model.computeIIS()

    # 获取不可行约束
    infeasible_constraints = []
    for c in model.getConstrs():
        if c.IISConstr:
            infeasible_constraints.append(c.ConstrName)

    # 打印不可行约束
    print(f"发现 {len(infeasible_constraints)} 个不可行约束:")
    for i, constr_name in enumerate(infeasible_constraints):
        print(f"{i + 1}. {constr_name}")

    return infeasible_constraints


def load_data():
    """加载所有需要的数据"""
    # 加载现状数据
    current_status = pd.read_excel(CURRENT_STATUS_FILE)
    current_status = current_status[(current_status['Date'] == 2) & (current_status['flight type'] == 'PAX')]

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
        if airline == "其它":
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


def complete_day1_assignments(day3_dep_assignments, arrival_flights, departure_flights):
    """
    用（DAY2进港-DAY3离港）航班信息，补齐（DAY1进港-DAY2离港航班信息）

    Args:
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
