import gurobipy as gp
from gurobipy import GRB
import pandas as pd

# 现状波形数据
original = [
    0, 1, 0, 0, 0, 0, 0,
    15, 22, 9, 7, 4, 6, 7,
    9, 4, 5, 6, 3, 6, 7,
    6, 0, 1
]

# 创建模型
model = gp.Model("Integer_Waveform")
x = model.addVars(24, lb=0, vtype=GRB.INTEGER, name="x")  # 整数变量

# 约束条件
model.addConstr(x.sum() == 228, "Total")

# 目标函数：最小化平方差
model.setObjective(gp.quicksum((x[i] - original[i]) ** 2 for i in range(24)), GRB.MINIMIZE)

# # 需要统计学计算（需线性化近似）
# mean_orig = sum(original)/24
# covariance = gp.quicksum((x[i] - original[i])*(original[i] - mean_orig) for i in range(24))
# model.setObjective(-covariance, GRB.MINIMIZE)  # 最大化协方差（负号实现最小化）

# 求解
model.optimize()

# 结果处理与保存
if model.status == GRB.OPTIMAL:
    future = [int(x[i].x) for i in range(24)]  # 确保整数结果

    # 创建对比表格
    df = pd.DataFrame({
        '小时': range(24),
        '现状': original,
        '未来': future,
        '增量': [future[i] - original[i] for i in range(24)],
        '平方差': [(future[i] - original[i]) ** 2 for i in range(24)]
    })

    # 添加汇总行
    df.loc[24] = ['总计', sum(original), sum(future), sum(df['增量']), sum(df['平方差'])]

    # 保存到Excel
    df.to_excel("波形对比.xlsx", index=False, engine='openpyxl')

    print("优化成功！文件已保存为 波形对比.xlsx")
    print(f"现状总和: {sum(original)}")
    print(f"未来总和: {sum(future)}")
else:
    print("未找到可行解")
