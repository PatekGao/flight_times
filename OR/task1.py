import gurobipy as gp
import numpy as np
import pandas as pd
from gurobipy import GRB
from OR.pre_task1 import ref_arr_dom, ref_dep_int, ref_dep_dom, ref_arr_int

H_arr_dom = [65, 22, 2, 2, 0, 0, 0, 2, 7, 36, 51, 55, 52, 55, 58, 52, 54, 45, 46, 53, 47, 47, 55, 60]
H_arr_int = [6, 7, 15, 3, 7, 13, 11, 5, 4, 4, 3, 9, 7, 3, 4, 9, 7, 9, 7, 7, 6, 10, 11, 11]
H_dep_dom = [2, 2, 3, 1, 0, 3, 56, 60, 70, 48, 48, 47, 52, 53, 40, 44, 40, 60, 52, 53, 55, 50, 17, 10]
H_dep_int = [4, 11, 11, 11, 5, 6, 2, 6, 7, 16, 11, 2, 6, 9, 11, 6, 13, 4, 11, 5, 7, 3, 3, 7]

# ref_arr = [42, 44, 40, 36, 35, 33, 31, 33, 29, 28, 25, 20, 15, 15, 15, 14, 16, 14, 12, 9, 8, 6, 6, 6, 7, 6, 5, 5, 3, 3,
#            3, 2, 2, 2, 2, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 4, 4, 4, 4, 4, 4, 5, 4, 4, 4, 4, 3,
#            3, 3, 3, 3, 2, 3, 4, 4, 4, 4, 4, 4, 3, 3, 3, 3, 3, 2, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 1, 1, 2, 3, 4, 4,
#            6, 6, 7, 10, 13, 16, 17, 18, 17, 18, 17, 17, 18, 18, 23, 26, 28, 28, 29, 31, 36, 35, 36, 36, 40, 41, 38, 34,
#            34, 37, 37, 34, 31, 33, 33, 34, 30, 32, 31, 32, 31, 29, 30, 30, 30, 30, 30, 30, 29, 28, 31, 35, 36, 38, 36,
#            40, 44, 44, 44, 46, 48, 49, 48, 43, 39, 37, 40, 41, 39, 36, 39, 38, 35, 34, 32, 32, 32, 34, 33, 34, 32, 35,
#            35, 33, 33, 32, 31, 31, 34, 31, 30, 28, 27, 27, 24, 26, 28, 30, 31, 32, 31, 35, 34, 31, 34, 34, 35, 36, 36,
#            33, 32, 34, 33, 28, 29, 29, 29, 28, 28, 26, 28, 30, 31, 28, 30, 31, 34, 34, 34, 36, 39, 40, 38, 37, 36, 36,
#            34, 35, 30, 33, 33, 31, 26, 25, 23, 26, 29, 31, 31, 30, 38, 36, 33, 34, 35, 34, 38, 36, 35, 35, 34, 35, 32,
#            35, 36, 36, 39, 40, 38, 40, 43, 41, 43, 40, 40, 35, 33, 30, 26, 24, 21, 17, 11, 9, 6, 5, 0]
# ref_dep = [3, 3, 2, 1, 2, 2, 2, 1, 1, 2, 2, 4, 4, 5, 5, 5, 4, 5, 5, 7, 7, 6, 6, 5, 5, 5, 6, 6, 7, 6, 6, 4, 4, 4, 4, 3,
#            3, 2, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 0, 0, 0, 0, 0, 1, 2,
#            6, 12, 14, 18, 20, 23, 29, 34, 39, 46, 52, 53, 51, 54, 55, 55, 60, 60, 60, 63, 62, 57, 50, 51, 50, 50, 53,
#            49, 50, 53, 49, 47, 46, 45, 47, 46, 47, 41, 39, 41, 33, 29, 28, 25, 25, 25, 27, 28, 29, 31, 29, 27, 29, 27,
#            26, 27, 27, 30, 29, 28, 30, 30, 33, 36, 37, 37, 38, 37, 38, 37, 35, 36, 34, 36, 34, 34, 34, 35, 36, 39, 37,
#            36, 35, 34, 37, 33, 30, 28, 26, 26, 26, 25, 23, 25, 28, 31, 28, 27, 28, 30, 32, 33, 34, 33, 34, 36, 36, 31,
#            30, 29, 31, 32, 33, 33, 31, 29, 31, 27, 27, 29, 29, 33, 33, 32, 30, 29, 29, 35, 36, 36, 34, 36, 37, 36, 36,
#            35, 34, 36, 40, 34, 35, 37, 38, 37, 37, 39, 39, 40, 42, 40, 35, 39, 36, 33, 32, 31, 30, 28, 28, 30, 33, 35,
#            36, 36, 36, 36, 35, 38, 38, 38, 37, 37, 37, 34, 34, 30, 31, 31, 33, 32, 30, 28, 27, 23, 18, 19, 20, 18, 16,
#            15, 14, 10, 11, 10, 10, 9, 8, 7, 5, 6, 5, 6, 5, 6, 8, 7, 6, 6, 6, 6, 6, 5, 5, 4, 4, 3, 0]

ARR_DOM_MAX = 71
ARR_INT_MAX = 16
DEP_DOM_MAX = 77
DEP_INT_MAX = 18

ARR_MAX = 78
DEP_MAX = 86
DOM_MAX = 119
INT_MAX = 26
TOT_MAX = 132

ARR_LIMIT = 12
DEP_LIMIT = 12
TOT_LIMIT = 20

# H_arr_dom, H_arr_int, H_dep_dom, H_dep_int 是长度为24的列表，表示每小时的总和
# ARR_DOM_MAX, ARR_INT_MAX, DEP_DOM_MAX, DEP_INT_MAX 对应各分类的动态小时最大值
# ARR_MAX, DEP_MAX, DOM_MAX, INT_MAX, TOT_MAX 是整合类别的动态小时最大值
# ARR_LIMIT, DEP_LIMIT, TOT_LIMIT 是每五分钟的最大值，假设是标量或数组


model = gp.Model('dynamic_schedule')

# 定义变量
num_periods = 24 * 12  # 288 five-minute periods
arr_dom = model.addVars(num_periods, vtype=GRB.INTEGER, name="arr_dom")
arr_int = model.addVars(num_periods, vtype=GRB.INTEGER, name="arr_int")
dep_dom = model.addVars(num_periods, vtype=GRB.INTEGER, name="dep_dom")
dep_int = model.addVars(num_periods, vtype=GRB.INTEGER, name="dep_int")
b_indicators = {
    'ARR': model.addVars(277, vtype=GRB.BINARY, name="b_ARR"),
    'DEP': model.addVars(277, vtype=GRB.BINARY, name="b_DEP"),
    'TOT': model.addVars(277, vtype=GRB.BINARY, name="b_TOT"),
    'ARR_DOM': model.addVars(277, vtype=GRB.BINARY, name="b_ARR_DOM"),
    'ARR_INT': model.addVars(277, vtype=GRB.BINARY, name="b_ARR_INT"),
    'DEP_DOM': model.addVars(277, vtype=GRB.BINARY, name="b_DEP_DOM"),
    'DEP_INT': model.addVars(277, vtype=GRB.BINARY, name="b_DEP_INT"),
    'DOM': model.addVars(277, vtype=GRB.BINARY, name="b_DOM"),
    'INT': model.addVars(277, vtype=GRB.BINARY, name="b_INT")
}

# 约束1：每小时的总和等于静态值
for h in range(24):
    start = h * 12
    end = start + 12
    model.addConstr(gp.quicksum(arr_dom[t] for t in range(start, end)) == H_arr_dom[h], name="H_ARR_DOM")
    model.addConstr(gp.quicksum(arr_int[t] for t in range(start, end)) == H_arr_int[h], name="H_ARR_INT")
    model.addConstr(gp.quicksum(dep_dom[t] for t in range(start, end)) == H_dep_dom[h], name="H_dep_dom")
    model.addConstr(gp.quicksum(dep_int[t] for t in range(start, end)) == H_dep_int[h], name="H_dep_int")

# 约束2：滑动窗口总和 <= 动态小时最大值（细分类别）
for i in range(277):  # 277滑动窗口
    window = range(i, i + 12)
    # 细分类别
    sum_arr_dom = gp.quicksum(arr_dom[t] for t in window)
    model.addConstr(sum_arr_dom <= ARR_DOM_MAX, name="ARR_DOM")
    sum_arr_int = gp.quicksum(arr_int[t] for t in window)
    model.addConstr(sum_arr_int <= ARR_INT_MAX, name="ARR_INT")
    sum_dep_dom = gp.quicksum(dep_dom[t] for t in window)
    model.addConstr(sum_dep_dom <= DEP_DOM_MAX, name="DEP_DOM")
    sum_dep_int = gp.quicksum(dep_int[t] for t in window)
    model.addConstr(sum_dep_int <= DEP_INT_MAX, name="DEP_INT")
    # 整合类别
    sum_ARR = sum_arr_dom + sum_arr_int
    model.addConstr(sum_ARR <= ARR_MAX, name="ARR_MAX")
    sum_DEP = sum_dep_dom + sum_dep_int
    model.addConstr(sum_DEP <= DEP_MAX, name="DEP_MAX")
    sum_DOM = sum_arr_dom + sum_dep_dom
    model.addConstr(sum_DOM <= DOM_MAX, name="DOM_MAX")
    sum_INT = sum_arr_int + sum_dep_int
    model.addConstr(sum_INT <= INT_MAX, name="INT_MAX")
    sum_TOT = sum_ARR + sum_DEP
    model.addConstr(sum_TOT <= TOT_MAX, name="TOT_MAX")

# 约束3： 动态小时最大值等于限制
for i in range(277):
    window = range(i, i + 12)

    # 计算各指标窗口总和
    sum_arr_dom = gp.quicksum(arr_dom[t] for t in window)
    sum_arr_int = gp.quicksum(arr_int[t] for t in window)
    sum_dep_dom = gp.quicksum(dep_dom[t] for t in window)
    sum_dep_int = gp.quicksum(dep_int[t] for t in window)

    # 添加指标约束（当二元变量=1时，必须达到MAX值）
    model.addConstr(sum_arr_dom >= ARR_DOM_MAX * b_indicators['ARR_DOM'][i], name="ARR_DOM=MAX")
    model.addConstr(sum_arr_int >= ARR_INT_MAX * b_indicators['ARR_INT'][i], name="ARR_INT=MAX")
    model.addConstr(sum_dep_dom >= DEP_DOM_MAX * b_indicators['DEP_DOM'][i], name="DEP_DOM=MAX")
    model.addConstr(sum_dep_int >= DEP_INT_MAX * b_indicators['DEP_INT'][i], name="DEP_INT=MAX")

    model.addConstr((sum_arr_dom + sum_arr_int) >= ARR_MAX * b_indicators['ARR'][i], name="ARR=MAX")
    model.addConstr((sum_dep_dom + sum_dep_int) >= DEP_MAX * b_indicators['DEP'][i], name="DEP=MAX")
    model.addConstr((sum_arr_dom + sum_dep_dom) >= DOM_MAX * b_indicators['DOM'][i], name="DOM=MAX")
    model.addConstr((sum_arr_int + sum_dep_int) >= INT_MAX * b_indicators['INT'][i], name="INT=MAX")
    model.addConstr((sum_arr_dom + sum_arr_int + sum_dep_dom + sum_dep_int) >= TOT_MAX * b_indicators['TOT'][i],
                    name="TOT=MAX")

# 添加至少一个窗口达标约束
for indicator in ['ARR', 'DEP', 'TOT', 'ARR_DOM', 'ARR_INT', 'DEP_DOM', 'DEP_INT', 'DOM', 'INT']:
    model.addConstr(gp.quicksum(b_indicators[indicator]) >= 1, name=indicator)

# 约束4：每五分钟的即时限制
for t in range(num_periods):
    # ARR = arr_dom + arr_int
    model.addConstr(arr_dom[t] + arr_int[t] <= ARR_LIMIT, name="ARR_LIMIT")
    # DEP = dep_dom + dep_int
    model.addConstr(dep_dom[t] + dep_int[t] <= DEP_LIMIT, name="DEP_LIMIT")
    # TOT = ARR + DEP
    model.addConstr((arr_dom[t] + arr_int[t] + dep_dom[t] + dep_int[t]) <= TOT_LIMIT, name="TOT_LIMIT")

# 约束5：每五分钟未来值 >= 现状值
for t in range(num_periods):
    model.addConstr(arr_dom[t] >= ref_arr_dom[t], name="REF_ARR_DOM")
    model.addConstr(dep_dom[t] >= ref_dep_dom[t], name="REF_DEP_DOM")
    model.addConstr(arr_int[t] >= ref_arr_int[t], name="REF_ARR_INT")
    model.addConstr(dep_int[t] >= ref_dep_int[t], name="REF_DEP_INT")

# 🆕 整数规划参数调优
model.Params.IntegralityFocus = 1  # 强调整数可行性

# model.Params.MIPGap = 0.1  # shape允许间隙
model.Params.MIPGap = 0.99  # smooth允许间隙

model.Params.Heuristics = 1  # 增加启发式搜索
model.Params.Presolve = 1  # 基础预处理

# === 波形优化目标 ===
# shape_obj = gp.QuadExpr()
# for t in range(num_periods):
#     # 计算总ARR和DEP
#     total_arr = arr_dom[t] + arr_int[t]
#     total_dep = dep_dom[t] + dep_int[t]
#
#     # 动态权重：增强早晚高峰匹配
#     weight = 1.0
#     if (7 * 12 <= t < 9 * 12) or (17 * 12 <= t < 19 * 12):
#         weight = 3.0  # 早晚高峰权重提升
#
#     # 平方差项
#     shape_obj += weight * (total_arr - ref_arr[t]) ** 2
#     shape_obj += weight * (total_dep - ref_dep[t]) ** 2

# === 平滑扰动优化目标 ===
np.random.seed(42)  # 可设置的随机种子
noise_weights = np.random.uniform(0.5, 1.5, size=(4, 287))  # 4个变量类型，287个间隔
smooth_obj = gp.QuadExpr()
arr_total = {t: arr_dom[t] + arr_int[t] for t in range(288)}
dep_total = {t: dep_dom[t] + dep_int[t] for t in range(288)}

for var_idx, var_list in enumerate([arr_total, dep_total]):
    for t in range(287):
        diff = var_list[t + 1] - var_list[t]
        # 核心修改：波动项权重引入随机性
        weight = noise_weights[var_idx, t]
        smooth_obj += weight * (diff * diff)  # 随机权重影响波动幅度
max_delta = 3  # 允许相邻时段最大变化量
for var_list in [arr_dom, arr_int, dep_dom, dep_int]:
    for t in range(287):
        model.addConstr(var_list[t + 1] - var_list[t] <= max_delta, name="max")
        model.addConstr(var_list[t + 1] - var_list[t] >= -max_delta, name="min")

for var_list in [arr_dom, arr_int, dep_dom, dep_int]:
    for t in range(287):
        diff = var_list[t + 1] - var_list[t]
        smooth_obj += diff * diff  # 仍使用二次项

# === 随机生成优化目标 ===

model.setObjective(smooth_obj, GRB.MINIMIZE)

# 求解模型
model.optimize()

# 提取结果
if model.status == GRB.OPTIMAL:
    dynamic_schedule = []
    for t in range(num_periods):
        arr_dom_val = arr_dom[t].X
        arr_int_val = arr_int[t].X
        dep_dom_val = dep_dom[t].X
        dep_int_val = dep_int[t].X
        dynamic_schedule.append((arr_dom_val, arr_int_val, dep_dom_val, dep_int_val))
    print(dynamic_schedule)
    results = [(arr_dom[t].X, arr_int[t].X, dep_dom[t].X, dep_int[t].X)
               for t in range(num_periods)]

    # 生成时间戳列表
    time_index = pd.date_range("00:00", "23:55", freq="5min").strftime("%H:%M")

    # 创建DataFrame
    df = pd.DataFrame(results,
                      columns=["ARR_DOM", "ARR_INT", "DEP_DOM", "DEP_INT"],
                      index=time_index)
    df.reset_index(inplace=True)
    df.rename(columns={"index": "Time"}, inplace=True)

    # 保存到Excel
    df.to_excel("dynamic_sheet.xlsx", index=False)

else:
    print("!!!!!!!! No solution found !!!!!!!!")
