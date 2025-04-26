# 航司航向匹配优化系统

## 项目概述

本项目是一个航班调度优化系统，主要用于解决机场航班的航司和航向分配问题。系统通过数学优化模型（整数线性规划），根据各种约束条件，为进港和离港航班分配最优的航司和航向组合，以满足机场运营需求。

## 系统架构

系统由以下几个主要模块组成：

1. **数据处理模块**：负责从Excel文件中读取输入数据，包括航司信息、航向信息、宽体机数据等。
2. **进港航班分配模块**：为进港航班分配航司和航向。
3. **离港航班分配模块**：为离港航班分配航司和航向，并与进港航班保持一致性。
4. **结果处理模块**：处理优化结果并输出到Excel文件。
5. **工具函数模块**：提供各种辅助功能。

系统的工作流程如下：

1. 从Excel模板读取输入数据
2. 加载现状数据和配对数据
3. 分配国内进港航班
4. 分配国际进港航班
5. 分配国内离港航班
6. 分配国际离港航班
7. 补齐日期为1的进港航班和日期为3的离港航班
8. 处理最终结果并输出

## 文件说明

### 主要文件

- `airline_heading_match.py`：主程序，协调各模块工作
- `excel_to_dataset.py`：从Excel模板读取输入数据
- `arrival_assignment.py`：进港航班分配模块
- `departure_assignment.py`：离港航班分配模块
- `process_final_result.py`：结果处理模块
- `utils.py`：工具函数模块
- `config.py`：配置文件

### 输入文件

- `输入模板.xlsx`：包含航司、航向等基础数据的Excel模板
- `2024_现状.xlsx`：现状数据
- `到达航班比例表.xlsx`：进港航班小时统计数据
- `出发航班比例表.xlsx`：离港航班小时统计数据
- `final_pairing_processed.xlsx`：航班配对数据

### 输出文件

- `final_result.xlsx`：优化结果
- `final_result_processed.xlsx`：处理后的最终结果

## 函数详解

### excel_to_dataset.py

#### read_excel_template()

读取Excel模板文件，提取所有必要的数据。

内部函数：

- `read_domestic_airlines(df)`：读取国内航班航司数据
- `read_international_airlines(df, start_row)`：读取国际航班航司数据
- `read_arrival_headings(df, start_row)`：读取进港航向数据
- `read_departure_headings(df, start_row)`：读取离港航向数据
- `read_wide_body_data(df, start_row)`：读取航司宽体机数据
- `read_absolute_long_routing(df, start_row)`：读取绝对远程航向分配航司集团数据
- `read_wide_body_exception_routing(df, start_row)`：读取不能有宽体机的航向数据
- `read_wide_body_ratio_increase_routing(df, start_row)`：读取宽体机比例升高的航向数据
- `read_main_heading_exception_airlines(df, start_row)`：读取主航向比例约束例外航司集团数据
- `read_wave_exception_airlines(df, start_row)`：读取波形约束例外航司集团数据

### arrival_assignment.py

#### assign_arrival_flights(current_status, arrival_flights, market_type, hourly_dom_stats, hourly_int_stats, departure_flights, prev_dom_dep_counts, prev_int_dep_counts)

为进港航班分配航司和航向。

参数：
- `current_status`：现状数据
- `arrival_flights`：需要分配的进港航班
- `market_type`：'DOM'（国内）或'INT'（国际）
- `hourly_dom_stats`：国内航班小时统计数据
- `hourly_int_stats`：国际航班小时统计数据
- `departure_flights`：所有离港航班数据，用于检查配额限制
- `prev_dom_dep_counts`：之前分配的国内离港航班数量
- `prev_int_dep_counts`：之前分配的国际离港航班数量

返回：
- 带有航司和航向分配的进港航班数据，以及各航司的DOM和INT离港航班分配数量

### departure_assignment.py

#### assign_departure_flights(arrival_assignments, departure_flights, current_status, market_type, hourly_dom_dep_stats, hourly_int_dep_stats)

为离港航班分配航司和航向。

参数：
- `arrival_assignments`：已分配的进港航班
- `departure_flights`：需要分配的离港航班
- `current_status`：现状数据
- `market_type`：'DOM'（国内）或'INT'（国际）
- `hourly_dom_dep_stats`：国内出发航班小时统计数据
- `hourly_int_dep_stats`：国际出发航班小时统计数据

返回：
- 带有航司和航向分配的离港航班数据

### utils.py

#### diagnose_infeasibility(model)

诊断Gurobi模型不可行的原因。

#### load_data()

加载所有需要的数据，包括现状数据、配对数据和小时统计数据。

#### get_hour_from_time(time_str)

从时间字符串中提取小时。

#### get_main_headings(current_status)

从现状数据中获取各航司的主航向。

#### complete_day1_assignments(day2_dep_assignments, arrival_flights, departure_flights)

补齐日期为1的进港航班分配。

#### complete_day3_assignments(day2_arr_assignments, departure_flights)

补齐日期为3的离港航班分配。

### process_final_result.py

#### process_excel(INPUT_FILE, OUTPUT_FILE)

处理优化结果，合并进港和离港航班数据，并输出到Excel文件。

### config.py

包含系统配置参数，如文件路径、MIP Gap、波形偏差和宽体机偏差等。

## 优化模型说明

系统使用Gurobi求解器实现整数线性规划模型，主要考虑以下约束条件：

1. **航司约束**：每个航司分配的航班数量需符合要求
2. **航向约束**：每个航向分配的航班数量需符合要求
3. **主航向约束**：保持各航司的主航向比例
4. **波形约束**：保持各小时的航班分布
5. **宽体机约束**：特定航向的宽体机分配规则
6. **绝对远程航向约束**：绝对远程航向只分配给特定航司
7. **航班配对约束**：保持进港和离港航班的一致性

## 使用方法

1. 准备输入数据文件
2. 配置`config.py`中的参数
3. 运行`airline_heading_match.py`主程序
4. 查看输出结果

## 注意事项

- 系统依赖Gurobi优化求解器，需要安装相应的Python包
- 输入数据格式需严格按照模板要求
- 优化过程可能需要较长时间，取决于问题规模和复杂度