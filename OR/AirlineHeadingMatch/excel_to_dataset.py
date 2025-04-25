import pandas as pd


def read_excel_template():
    # 读取Excel文件
    file_path = '输入模板.xlsx'

    def read_domestic_airlines(df):
        # 从第4行开始读取数据
        start_row = 2  # pandas的索引从0开始，所以第4行是索引3
        DOM_AIRLINES = {'主基地': {}, '非主基地': {}}

        current_row = start_row
        # 读取直到遇到第一列为空的行
        while pd.notna(df.iloc[current_row, 0]):
            airline_type = df.iloc[current_row, 0]  # 主基地/非主基地
            airline_group = df.iloc[current_row, 1]  # 航司集团名称
            arr_flights = df.iloc[current_row, 2]  # 到达航班数
            dep_flights = df.iloc[current_row, 3]  # 出发航班数

            # 将数据添加到字典中
            if airline_type in DOM_AIRLINES:
                DOM_AIRLINES[airline_type][airline_group] = {
                    'ARR': int(arr_flights),
                    'DEP': int(dep_flights)
                }

            current_row += 1

        return DOM_AIRLINES

    def read_international_airlines(df, start_row):
        INT_AIRLINES = {'主基地': {}, '非主基地': {}}

        current_row = start_row
        # 读取直到遇到第一列为空的行
        while pd.notna(df.iloc[current_row, 0]):
            airline_type = df.iloc[current_row, 0]  # 主基地/非主基地
            airline_group = df.iloc[current_row, 1]  # 航司集团名称
            arr_flights = df.iloc[current_row, 2]  # 到达航班数
            dep_flights = df.iloc[current_row, 3]  # 出发航班数

            # 将数据添加到字典中
            if airline_type in INT_AIRLINES:
                INT_AIRLINES[airline_type][airline_group] = {
                    'ARR': int(arr_flights),
                    'DEP': int(dep_flights)
                }

            current_row += 1

        return INT_AIRLINES

    def read_arrival_headings(df, start_row):
        HEADINGS_ARR = {}

        current_row = start_row
        # 读取直到遇到第一列为空的行
        while pd.notna(df.iloc[current_row, 0]):
            try:
                heading = df.iloc[current_row, 0]  # 航向名称
                dom_flights = df.iloc[current_row, 1]  # 国内航班数
                int_flights = df.iloc[current_row, 2]  # 国际航班数
                int_type = df.iloc[current_row, 3]  # 国际性质

                # 跳过标题行
                if isinstance(heading, str) and ('DOM' in str(dom_flights) or '国内' in str(dom_flights)):
                    current_row += 1
                    continue

                # 将数据添加到字典中
                HEADINGS_ARR[heading] = {
                    'DOM': int(float(dom_flights)) if pd.notna(dom_flights) else 0,
                    'INT': int(float(int_flights)) if pd.notna(int_flights) else 0,
                    'INT性质': str(int_type) if pd.notna(int_type) else ''
                }
            except (ValueError, TypeError) as e:
                print(f"警告：在第 {current_row + 1} 行出现数据格式问题，已跳过该行")
                print(f"详细信息：{e}")

            current_row += 1

        return HEADINGS_ARR

    def read_departure_headings(df, start_row):
        HEADINGS_DEP = {}

        current_row = start_row
        # 读取直到遇到第一列为空的行
        while pd.notna(df.iloc[current_row, 0]):
            try:
                heading = df.iloc[current_row, 0]  # 航向名称
                dom_flights = df.iloc[current_row, 1]  # 国内航班数
                int_flights = df.iloc[current_row, 2]  # 国际航班数
                int_type = df.iloc[current_row, 3]  # 国际性质

                # 跳过标题行
                if isinstance(heading, str) and ('DOM' in str(dom_flights) or '国内' in str(dom_flights)):
                    current_row += 1
                    continue

                # 将数据添加到字典中
                HEADINGS_DEP[heading] = {
                    'DOM': int(float(dom_flights)) if pd.notna(dom_flights) else 0,
                    'INT': int(float(int_flights)) if pd.notna(int_flights) else 0,
                    'INT性质': str(int_type) if pd.notna(int_type) else ''
                }
            except (ValueError, TypeError) as e:
                print(f"警告：在第 {current_row + 1} 行出现数据格式问题，已跳过该行")
                print(f"详细信息：{e}")

            current_row += 1

        return HEADINGS_DEP

    def read_wide_body_data(df, start_row):
        AIRLINES_WIDE = {}

        current_row = start_row
        # 读取直到遇到第一列为空的行
        while pd.notna(df.iloc[current_row, 0]):
            try:
                airline_group = df.iloc[current_row, 0]  # 航司集团名称
                dom_flights = df.iloc[current_row, 1]  # 国内航班数
                int_flights = df.iloc[current_row, 2]  # 国际航班数

                # 跳过标题行
                if isinstance(airline_group, str) and ('DOM' in str(dom_flights) or '国内' in str(dom_flights)):
                    current_row += 1
                    continue

                # 将数据添加到字典中
                AIRLINES_WIDE[airline_group] = {
                    'DOM': int(float(dom_flights)) if pd.notna(dom_flights) else 0,
                    'INT': int(float(int_flights)) if pd.notna(int_flights) else 0
                }
            except (ValueError, TypeError) as e:
                print(f"警告：在第 {current_row + 1} 行出现数据格式问题，已跳过该行")
                print(f"详细信息：{e}")

            current_row += 1

        return AIRLINES_WIDE

    def read_absolute_long_routing(df, start_row):
        ABSOLUTE_LONG_ROUTING = []

        # 读取start_row + 1行的数据（跳过标题行）
        current_row = start_row - 1
        current_col = 1  # 从第2列开始读取

        # 读取直到遇到空白单元格
        while current_col < len(df.columns) and pd.notna(df.iloc[current_row, current_col]):
            airline_group = df.iloc[current_row, current_col]
            ABSOLUTE_LONG_ROUTING.append(str(airline_group))
            current_col += 1

        return ABSOLUTE_LONG_ROUTING

    def read_wide_body_exception_routing(df, start_row):
        # 初始化四个列表
        ARR_DOM_WIDE_EXCEPTION_ROUTING = []  # 进港国内
        ARR_INT_WIDE_EXCEPTION_ROUTING = []  # 进港国际
        DEP_DOM_WIDE_EXCEPTION_ROUTING = []  # 离港国内
        DEP_INT_WIDE_EXCEPTION_ROUTING = []  # 离港国际

        try:
            # 读取四行数据，每行从第3列开始
            for i in range(4):
                current_row = start_row - 1 + i  # 因为之前加了2，这里需要减回去
                current_col = 2  # 从第3列开始读取

                # 读取直到遇到空白单元格
                while current_col < len(df.columns) and pd.notna(df.iloc[current_row, current_col]):
                    heading = str(df.iloc[current_row, current_col])
                    # 根据行号将数据添加到对应的列表中
                    if i == 0:
                        ARR_DOM_WIDE_EXCEPTION_ROUTING.append(heading)
                    elif i == 1:
                        ARR_INT_WIDE_EXCEPTION_ROUTING.append(heading)
                    elif i == 2:
                        DEP_DOM_WIDE_EXCEPTION_ROUTING.append(heading)
                    else:
                        DEP_INT_WIDE_EXCEPTION_ROUTING.append(heading)
                    current_col += 1
        except Exception as e:
            print(f"读取不能有宽体机的航向数据时出错：{str(e)}")

        return (ARR_DOM_WIDE_EXCEPTION_ROUTING,
                ARR_INT_WIDE_EXCEPTION_ROUTING,
                DEP_DOM_WIDE_EXCEPTION_ROUTING,
                DEP_INT_WIDE_EXCEPTION_ROUTING)

    def read_wide_body_ratio_increase_routing(df, start_row):
        # 初始化四个列表
        ARR_DOM_WIDE_RATIO_INCREASE = []  # 进港国内
        ARR_INT_WIDE_RATIO_INCREASE = []  # 进港国际
        DEP_DOM_WIDE_RATIO_INCREASE = []  # 离港国内
        DEP_INT_WIDE_RATIO_INCREASE = []  # 离港国际

        try:
            # 读取四行数据，每行从第3列开始
            for i in range(4):
                current_row = start_row - 1 + i  # 因为之前加了2，这里需要减回去
                current_col = 2  # 从第3列开始读取

                # 读取直到遇到空白单元格
                while current_col < len(df.columns) and pd.notna(df.iloc[current_row, current_col]):
                    heading = str(df.iloc[current_row, current_col])
                    # 根据行号将数据添加到对应的列表中
                    if i == 0:
                        ARR_DOM_WIDE_RATIO_INCREASE.append(heading)
                    elif i == 1:
                        ARR_INT_WIDE_RATIO_INCREASE.append(heading)
                    elif i == 2:
                        DEP_DOM_WIDE_RATIO_INCREASE.append(heading)
                    else:
                        DEP_INT_WIDE_RATIO_INCREASE.append(heading)
                    current_col += 1
        except Exception as e:
            print(f"读取宽体机比例升高的航向数据时出错：{str(e)}")

        return (ARR_DOM_WIDE_RATIO_INCREASE,
                ARR_INT_WIDE_RATIO_INCREASE,
                DEP_DOM_WIDE_RATIO_INCREASE,
                DEP_INT_WIDE_RATIO_INCREASE)

    def read_main_heading_exception_airlines(df, start_row):
        MAIN_HEADING_EXCEPTION_AIRLINES = []

        try:
            # 读取start_row行的数据
            current_row = start_row - 1
            current_col = 0  # 从第1列开始读取

            # 读取直到遇到空白单元格
            while current_col < len(df.columns) and pd.notna(df.iloc[current_row, current_col]):
                airline_group = df.iloc[current_row, current_col]
                MAIN_HEADING_EXCEPTION_AIRLINES.append(str(airline_group))
                current_col += 1
        except Exception as e:
            print(f"读取主航向比例约束例外航司集团数据时出错：{str(e)}")

        return MAIN_HEADING_EXCEPTION_AIRLINES

    def read_wave_exception_airlines(df, start_row):
        WAVE_EXCEPTION_AIRLINES = []

        try:
            # 读取start_row行的数据
            current_row = start_row - 1
            current_col = 0  # 从第1列开始读取

            # 读取直到遇到空白单元格
            while current_col < len(df.columns) and pd.notna(df.iloc[current_row, current_col]):
                airline_group = df.iloc[current_row, current_col]
                WAVE_EXCEPTION_AIRLINES.append(str(airline_group))
                current_col += 1
        except Exception as e:
            print(f"读取波形约束例外航司集团数据时出错：{str(e)}")

        return WAVE_EXCEPTION_AIRLINES

    try:
        # 读取输入模板sheet
        df = pd.read_excel(file_path, sheet_name='输入模板')

        # 需要查找的关键字列表
        keywords = [
            "2.国际航班航司数据",
            "3.进港航向数据",
            "4.离港航向数据",
            "5.航司宽体机数据",
            "6.绝对远程航向分配航司集团",
            "7.不能有宽体机的航向",
            "8.宽体机比例升高的航向",
            "9.主航向比例约束例外航司集团",
            "10.波形约束例外航司集团"
        ]

        # 查找每个关键字在第一列中的行号
        result = {}
        for keyword in keywords:
            # 查找关键字所在的行号（索引从0开始，所以加1得到实际行号）
            row_index = df.iloc[:, 0].str.strip().eq(keyword).idxmax() + 2
            result[keyword] = row_index

        print("\n各数据项所在行号：")
        for keyword, row in result.items():
            print(f"{keyword}: 第 {row} 行")

        # 读取国内航班航司数据
        domestic_airlines = read_domestic_airlines(df)
        print("\n国内航班航司数据:")
        print(domestic_airlines)

        # 读取国际航班航司数据
        international_airlines = read_international_airlines(df, result["2.国际航班航司数据"])
        print("\n国际航班航司数据:")
        print(international_airlines)

        # 读取进港航向数据
        arrival_headings = read_arrival_headings(df, result["3.进港航向数据"])
        print("\n进港航向数据:")
        print(arrival_headings)

        # 读取离港航向数据
        departure_headings = read_departure_headings(df, result["4.离港航向数据"])
        print("\n离港航向数据:")
        print(departure_headings)

        # 读取航司宽体机数据
        airlines_wide = read_wide_body_data(df, result["5.航司宽体机数据"])
        print("\n航司宽体机数据:")
        print(airlines_wide)

        # 读取绝对远程航向航司集团数据
        absolute_long_routing = read_absolute_long_routing(df, result["6.绝对远程航向分配航司集团"])
        print("\n绝对远程航向航司集团数据:")
        print(absolute_long_routing)

        # 读取不能有宽体机的航向数据
        (arr_dom_wide_exception, arr_int_wide_exception,
         dep_dom_wide_exception, dep_int_wide_exception) = read_wide_body_exception_routing(df, result[
            "7.不能有宽体机的航向"])
        print("\n不能有宽体机的航向数据:")
        print(f"进港国内：{arr_dom_wide_exception}")
        print(f"进港国际：{arr_int_wide_exception}")
        print(f"离港国内：{dep_dom_wide_exception}")
        print(f"离港国际：{dep_int_wide_exception}")

        # 读取宽体机比例升高的航向数据
        (arr_dom_wide_ratio_increase, arr_int_wide_ratio_increase,
         dep_dom_wide_ratio_increase, dep_int_wide_ratio_increase) = read_wide_body_ratio_increase_routing(df, result[
            "8.宽体机比例升高的航向"])
        print("\n宽体机比例升高的航向数据:")
        print(f"进港国内：{arr_dom_wide_ratio_increase}")
        print(f"进港国际：{arr_int_wide_ratio_increase}")
        print(f"离港国内：{dep_dom_wide_ratio_increase}")
        print(f"离港国际：{dep_int_wide_ratio_increase}")

        # 读取主航向比例约束例外航司集团数据
        main_heading_exception_airlines = read_main_heading_exception_airlines(df,
                                                                               result["9.主航向比例约束例外航司集团"])
        print("\n主航向比例约束例外航司集团数据:")
        print(main_heading_exception_airlines)

        # 读取波形约束例外航司集团数据
        wave_exception_airlines = read_wave_exception_airlines(df, result["10.波形约束例外航司集团"])
        print("\n波形约束例外航司集团数据:")
        print(wave_exception_airlines)

        return (result, df, domestic_airlines, international_airlines,
                arrival_headings, departure_headings, airlines_wide,
                absolute_long_routing,
                arr_dom_wide_exception, arr_int_wide_exception,
                dep_dom_wide_exception, dep_int_wide_exception,
                arr_dom_wide_ratio_increase, arr_int_wide_ratio_increase,
                dep_dom_wide_ratio_increase, dep_int_wide_ratio_increase,
                main_heading_exception_airlines,
                wave_exception_airlines)
    except Exception as e:
        print(f"读取Excel文件时发生错误: {str(e)}")
        return None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None


(row_indices, data, DOM_AIRLINES, INT_AIRLINES, HEADINGS_ARR, HEADINGS_DEP, AIRLINES_WIDE, ABSOLUTE_LONG_ROUTING,
 ARR_DOM_WIDE_EXCEPTION_ROUTING, ARR_INT_WIDE_EXCEPTION_ROUTING, DEP_DOM_WIDE_EXCEPTION_ROUTING,
 DEP_INT_WIDE_EXCEPTION_ROUTING, ARR_DOM_WIDE_UP_ROUTING, ARR_INT_WIDE_UP_ROUTING, DEP_DOM_WIDE_UP_ROUTING,
 DEP_INT_WIDE_UP_ROUTING, MAIN_HEADING_EXCEPTION_AIRLINES, WAVE_EXCEPTION_AIRLINES) = read_excel_template()
