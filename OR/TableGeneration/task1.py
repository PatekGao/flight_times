import gurobipy as gp
import numpy as np
import pandas as pd
from gurobipy import GRB

from utils import ref_arr_dom, ref_dep_int, ref_dep_dom, ref_arr_int, H_arr_dom, H_arr_int, H_dep_dom, H_dep_int, \
    ARR_DOM_MAX, ARR_INT_MAX, DEP_DOM_MAX, DEP_INT_MAX, ARR_MAX, DEP_MAX, DOM_MAX, INT_MAX, TOT_MAX, ARR_LIMIT, \
    DEP_LIMIT, TOT_LIMIT, ARR_15_LIMIT, DEP_15_LIMIT, TOT_15_LIMIT

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
    model.addConstr(gp.quicksum(arr_dom[t] for t in range(start, end)) == H_arr_dom[h], name=f"{h}小时_H_ARR_DOM")
    model.addConstr(gp.quicksum(arr_int[t] for t in range(start, end)) == H_arr_int[h], name=f"{h}小时_H_ARR_INT")
    model.addConstr(gp.quicksum(dep_dom[t] for t in range(start, end)) == H_dep_dom[h], name=f"{h}小时_H_dep_dom")
    model.addConstr(gp.quicksum(dep_int[t] for t in range(start, end)) == H_dep_int[h], name=f"{h}小时_H_dep_int")

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
    model.addConstr(arr_dom[t] >= ref_arr_dom[t], name=f"{5*t}分钟_REF_ARR_DOM")
    model.addConstr(dep_dom[t] >= ref_dep_dom[t], name=f"{5*t}分钟_REF_DEP_DOM")
    model.addConstr(arr_int[t] >= ref_arr_int[t], name=f"{5*t}分钟_REF_ARR_INT")
    model.addConstr(dep_int[t] >= ref_dep_int[t], name=f"{5*t}分钟_REF_DEP_INT")

# 约束6：进出港15分钟上限值 （进出港15分钟上限值均为28，双向为45）
for i in range(num_periods - 2):
    current_window = [i, i + 1, i + 2]

    # 进港总量 = 到达国内 + 到达国际
    arr_total = gp.quicksum(arr_dom[t] + arr_int[t] for t in current_window)
    # 出港总量 = 出发国内 + 出发国际
    dep_total = gp.quicksum(dep_dom[t] + dep_int[t] for t in current_window)
    # 双向总流量
    total = arr_total + dep_total

    model.addConstr(arr_total <= ARR_15_LIMIT, name=f"arr_15min_{i}")
    model.addConstr(dep_total <= DEP_15_LIMIT, name=f"dep_15min_{i}")
    model.addConstr(total <= TOT_15_LIMIT, name=f"total_15min_{i}")

# 🆕 整数规划参数调优
model.Params.IntegralityFocus = 1  # 强调整数可行性
model.Params.Heuristics = 1  # 增加启发式搜索
model.Params.Presolve = 1  # 基础预处理

# === 波形优化目标 ===
# model.Params.MIPGap = 0.1  # shape允许间隙
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
model.Params.MIPGap = 0.99  # smooth允许间隙
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

model.setObjective(smooth_obj, GRB.MINIMIZE)

# === 随机生成优化目标 ===
# model.Params.Seed = random.randint(0, 1000)

# 求解模型
model.optimize()

# 检查模型是否无解，如果无解则进行IIS分析
if model.status == 4:
    print("模型无解，正在分析导致无解的约束条件...")
    # 计算IIS（Irreducible Inconsistent Subsystem）
    model.computeIIS()
    print("\n以下约束条件导致模型无解:")
    for c in model.getConstrs():
        if c.IISConstr:
            print(f"约束名称: {c.ConstrName}")
            print(f"约束表达式: {c.Sense} {c.RHS}")
            print("-" * 50)

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
