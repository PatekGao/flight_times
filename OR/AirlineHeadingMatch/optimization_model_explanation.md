# 航司航向匹配优化系统数学模型详解

本文档详细解释了航司航向匹配优化系统中的数学优化模型，主要分析`arrival_assignment.py`和`departure_assignment.py`中的整数线性规划模型，包括决策变量、约束条件和目标函数，以及它们之间的联系。

## 1. 进港航班分配模型（arrival_assignment.py）

### 1.1 决策变量

进港航班分配模型使用三维决策变量 $x_{i,a,h}$，其中：
- $i$ 表示航班ID
- $a$ 表示航司
- $h$ 表示航向

决策变量 $x_{i,a,h}$ 是一个二元变量（0-1变量），当航班 $i$ 分配给航司 $a$ 和航向 $h$ 时，$x_{i,a,h} = 1$；否则，$x_{i,a,h} = 0$。

```python
x[flight_id, airline, heading] = model.addVar(vtype=GRB.BINARY, name=f"x_{flight_id}_{airline}_{heading}")
```

### 1.2 约束条件

#### 1.2.1 基本分配约束

每个航班只能分配给一个航司和一个航向：

$$\sum_{a \in A} \sum_{h \in H} x_{i,a,h} = 1, \forall i \in I$$

```python
model.addConstr(
    gp.quicksum(x[flight_id, airline, heading]
                for airline, _ in all_airlines
                for heading, _ in all_headings
                if (flight_id, airline, heading) in x) == 1,
    f"one_assignment_{flight_id}"
)
```

#### 1.2.2 航司配额约束

每个航司分配的航班数量必须等于其配额：

$$\sum_{i \in I} \sum_{h \in H} x_{i,a,h} = quota_a, \forall a \in A$$

```python
model.addConstr(
    gp.quicksum(x[flight_id, airline, heading]
                for i, flight in day2_arrivals.iterrows()
                for heading, _ in all_headings
                if (flight_id := flight['ID'], airline, heading) in x) == quota,
    f"airline_quota_{airline}"
)
```

#### 1.2.3 航向配额约束

每个航向分配的航班数量必须等于其配额：

$$\sum_{i \in I} \sum_{a \in A} x_{i,a,h} = quota_h, \forall h \in H$$

```python
model.addConstr(
    gp.quicksum(x[flight_id, airline, heading]
                for i, flight in day2_arrivals.iterrows()
                for airline, _ in all_airlines
                if (flight_id := flight['ID'], airline, heading) in x) == quota,
    f"heading_quota_{heading}"
)
```

#### 1.2.4 主航向比例约束

确保各航司的主航向比例不低于现状：

$$\sum_{i \in I} x_{i,a,h_{main}} \geq ratio_a \cdot \sum_{i \in I} \sum_{h \in H} x_{i,a,h}, \forall a \in A$$

其中，$h_{main}$ 是航司 $a$ 的主航向，$ratio_a$ 是现状中该航司主航向的比例。

```python
model.addConstr(
    main_heading_flights_expr >= current_ratio * total_flights_expr,
    f"main_heading_ratio_{airline}_{market_type}_Arrival"
)
```

此外，还确保主航向仍然是航班量最多的航向：

$$\sum_{i \in I} x_{i,a,h_{main}} \geq \sum_{i \in I} x_{i,a,h}, \forall a \in A, \forall h \in H, h \neq h_{main}$$

```python
model.addConstr(
    main_heading_flights_expr >= heading_flights_expr,
    f"main_heading_dominance_{airline}_{market_type}_{heading}_Arrival"
)
```

#### 1.2.5 离港航班配额约束

确保分配给各航司的进港航班对应的离港航班数量不超过配额：

$$\sum_{i \in I_{dom}} \sum_{h \in H} x_{i,a,h} + prev\_dom\_count_a \leq dom\_quota_a, \forall a \in A$$
$$\sum_{i \in I_{int}} \sum_{h \in H} x_{i,a,h} + prev\_int\_count_a \leq int\_quota_a, \forall a \in A$$

其中，$I_{dom}$ 和 $I_{int}$ 分别表示对应国内和国际离港航班的进港航班集合。

```python
model.addConstr(dom_dep_expr + prev_dom_count <= dom_quota, f"dom_dep_quota_{airline}")
model.addConstr(int_dep_expr + prev_int_count <= int_quota, f"int_dep_quota_{airline}")
```

#### 1.2.6 波形约束

确保各航司各小时的航班分布不低于现状：

$$\sum_{i \in I_h} \sum_{h \in H} x_{i,a,h} \geq current\_count_{a,h} - bias, \forall a \in A, \forall h \in \{0,1,...,23\}$$

其中，$I_h$ 表示在小时 $h$ 的航班集合，$current\_count_{a,h}$ 表示现状中航司 $a$ 在小时 $h$ 的航班数量，$bias$ 是允许的偏差。

```python
model.addConstr(
    future_count_expr >= current_count - bias,
    f"hourly_wave_{airline}_{hour}_{market_type}_Arrival"
)
```

#### 1.2.7 宽体机约束

确保各航司的宽体机航班数量在配额范围内：

$$wide\_body\_quota_a - bias \leq \sum_{i \in I_{wide}} \sum_{h \in H} x_{i,a,h} \leq wide\_body\_quota_a + bias, \forall a \in A$$

其中，$I_{wide}$ 表示宽体机航班集合。

```python
model.addConstr(
    wide_body_expr >= wide_body_quota - bias,
    f"wide_body_quota_{airline}_{market_type}_min_Arrival"
)
model.addConstr(
    wide_body_expr <= wide_body_quota + bias,
    f"wide_body_quota_{airline}_{market_type}_max_Arrival"
)
```

#### 1.2.8 宽体机比例升高约束

对于需要提高宽体机比例的航向，确保未来宽体机比例不低于现状：

$$\frac{\sum_{i \in I_{wide}} \sum_{a \in A} x_{i,a,h}}{\sum_{i \in I} \sum_{a \in A} x_{i,a,h}} \geq current\_wide\_ratio_h, \forall h \in H_{wide\_up}$$

其中，$H_{wide\_up}$ 表示需要提高宽体机比例的航向集合，$current\_wide\_ratio_h$ 表示现状中航向 $h$ 的宽体机比例。

```python
model.addConstr(
    future_wide_expr >= current_wide_ratio[heading] * future_total_expr,
    f"wide_body_ratio_increase_{heading}_{market_type}_Arrival"
)
```

#### 1.2.9 特殊约束

1. 绝对远程航向只分配给主基地航司或特定航司集团
2. 绝对远程航向必须分配宽体机
3. 特定航向不能分配宽体机
4. 航司只能分配给在现状中存在的航向

### 1.3 目标函数

进港航班分配模型使用多目标优化，包括两个主要目标：

1. 最小化各航司集团未来波形与现状波形的差异：

$$\sum_{a \in A} \sum_{h \in \{0,1,...,23\}} (future\_count_{a,h} - current\_count_{a,h})^2$$

2. 最小化未来分布与当前分布的偏离度：

$$\sum_{a \in A} \sum_{h \in H} \sum_{c \in C} \sum_{t \in \{0,1,...,23\}} |future\_distribution_{a,h,c,t} - current\_distribution_{a,h,c,t}|$$

其中，$C$ 表示机型类别（窄体机和宽体机）。

```python
model.setObjective(airline_wave_deviation * 10000 + obj_expr, GRB.MINIMIZE)
```

## 2. 离港航班分配模型（departure_assignment.py）

### 2.1 决策变量

离港航班分配模型使用两类主要决策变量：

1. 对于已分配航司的航班（与进港航班配对）：
   - $y_{i,h}$：航班 $i$ 分配给航向 $h$

2. 对于未分配航司的航班（进港航班日期为1的情况）：
   - $z_{i,a}$：航班 $i$ 分配给航司 $a$
   - $y_{i,h,a}$：航班 $i$ 分配给航向 $h$ 和航司 $a$

所有决策变量都是二元变量（0-1变量）。

```python
# 已分配航司的航班
y[flight_id, heading] = model.addVar(vtype=GRB.BINARY, name=f"y_{flight_id}_{heading}")

# 未分配航司的航班
z[flight_id, airline] = model.addVar(vtype=GRB.BINARY, name=f"z_{flight_id}_{airline}")
y[flight_id, heading, airline] = model.addVar(vtype=GRB.BINARY, name=f"w_{flight_id}_{airline}_{heading}")
```

### 2.2 约束条件

#### 2.2.1 基本分配约束

每个航班只能分配给一个航向：

对于已分配航司的航班：
$$\sum_{h \in H} y_{i,h} = 1, \forall i \in I_{assigned}$$

对于未分配航司的航班：
$$\sum_{a \in A} z_{i,a} = 1, \forall i \in I_{unassigned}$$
$$\sum_{h \in H} y_{i,h,a} = z_{i,a}, \forall i \in I_{unassigned}, \forall a \in A$$

```python
# 已分配航司的航班
model.addConstr(
    gp.quicksum(y[flight_id, heading]
                for heading, _ in all_headings
                if (flight_id, heading) in y) == 1,
    f"one_heading_{flight_id}"
)

# 未分配航司的航班
model.addConstr(
    gp.quicksum(z[flight_id, airline]
                for airline, _ in all_airlines
                if (flight_id, airline) in z) == 1,
    f"one_airline_{flight_id}"
)

model.addConstr(
    gp.quicksum(y[flight_id, heading, airline]
                for heading, _ in all_headings
                if (flight_id, heading, airline) in y) == z[flight_id, airline],
    f"one_heading_{flight_id}_{airline}"
)
```

#### 2.2.2 航司配额约束

每个航司分配的航班数量必须等于其配额：

$$assigned\_count_a + \sum_{i \in I_{unassigned}} z_{i,a} = quota_a, \forall a \in A$$

其中，$assigned\_count_a$ 表示已分配给航司 $a$ 的航班数量。

```python
model.addConstr(
    gp.quicksum(z[flight_id, airline]
                for i, flight in day2_departures.iterrows()
                if (flight_id := flight['ID'], airline) in z) == remaining_quota,
    f"airline_quota_{airline}"
)
```

#### 2.2.3 航向配额约束

每个航向分配的航班数量必须等于其配额：

$$\sum_{i \in I_{assigned}} y_{i,h} + \sum_{i \in I_{unassigned}} \sum_{a \in A} y_{i,h,a} = quota_h, \forall h \in H$$

```python
model.addConstr(assigned_expr + unassigned_expr == quota, f"heading_quota_{heading}")
```

#### 2.2.4 主航向比例约束

确保各航司的主航向比例不低于现状：

$$\frac{assigned\_main\_heading_a + \sum_{i \in I_{unassigned}} y_{i,h_{main},a}}{total\_assigned_a + \sum_{i \in I_{unassigned}} z_{i,a}} \geq ratio_a, \forall a \in A$$

其中，$h_{main}$ 是航司 $a$ 的主航向，$ratio_a$ 是现状中该航司主航向的比例，$assigned\_main\_heading_a$ 表示已分配给航司 $a$ 的主航向航班数量，$total\_assigned_a$ 表示已分配给航司 $a$ 的总航班数量。

```python
model.addConstr(
    (assigned_main_heading_expr + unassigned_main_heading_expr) >=
    current_ratio * (total_assigned_expr + total_unassigned_expr),
    f"main_heading_ratio_{airline}_{market_type}_Departure"
)
```

此外，还确保主航向仍然是航班量最多的航向：

$$assigned\_main\_heading_a + \sum_{i \in I_{unassigned}} y_{i,h_{main},a} \geq assigned\_heading_{a,h} + \sum_{i \in I_{unassigned}} y_{i,h,a}, \forall a \in A, \forall h \in H, h \neq h_{main}$$

```python
model.addConstr(
    (assigned_main_heading_expr + unassigned_main_heading_expr) >=
    (assigned_heading_expr + unassigned_heading_expr),
    f"main_heading_dominance_{airline}_{market_type}_{heading}_Departure"
)
```

#### 2.2.5 波形约束

确保各航司各小时的航班分布不低于现状：

$$assigned\_count_{a,t} + \sum_{i \in I_{unassigned,t}} \sum_{h \in H} y_{i,h,a} \geq current\_count_{a,t} - bias, \forall a \in A, \forall t \in \{0,1,...,23\}$$

其中，$I_{unassigned,t}$ 表示在小时 $t$ 的未分配航司航班集合，$assigned\_count_{a,t}$ 表示已分配给航司 $a$ 在小时 $t$ 的航班数量，$current\_count_{a,t}$ 表示现状中航司 $a$ 在小时 $t$ 的航班数量，$bias$ 是允许的偏差。

```python
model.addConstr(
    assigned_count_expr + unassigned_count_expr >= current_count - bias,
    f"hourly_wave_{airline}_{hour}_{market_type}_Departure"
)
```

#### 2.2.6 宽体机约束

确保各航司的宽体机航班数量在配额范围内：

$$wide\_body\_quota_a - bias \leq assigned\_wide\_body\_count_a + \sum_{i \in I_{unassigned,wide}} z_{i,a} \leq wide\_body\_quota_a + bias, \forall a \in A$$

其中，$I_{unassigned,wide}$ 表示未分配航司的宽体机航班集合，$assigned\_wide\_body\_count_a$ 表示已分配给航司 $a$ 的宽体机航班数量。

```python
model.addConstr(
    assigned_wide_body_count + unassigned_wide_body_expr >= wide_body_quota - bias,
    f"wide_body_quota_{airline}_{market_type}_min_Departure"
)
model.addConstr(
    assigned_wide_body_count + unassigned_wide_body_expr <= wide_body_quota + bias,
    f"wide_body_quota_{airline}_{market_type}_max_Departure"
)
```

#### 2.2.7 宽体机比例升高约束

对于需要提高宽体机比例的航向，确保未来宽体机比例不低于现状：

$$\frac{assigned\_wide\_expr_h + unassigned\_wide\_expr_h}{assigned\_total\_expr_h + unassigned\_total\_expr_h} \geq current\_wide\_ratio_h, \forall h \in H_{wide\_up}$$

其中，$H_{wide\_up}$ 表示需要提高宽体机比例的航向集合，$current\_wide\_ratio_h$ 表示现状中航向 $h$ 的宽体机比例。

```python
model.addConstr(
    wide_expr - 0.0001 >= current_wide_ratio[heading] * total_expr,
    f"wide_body_ratio_increase_{heading}_{market_type}_Departure"
)
```

#### 2.2.8 特殊约束

与进港航班分配模型类似，离港航班分配模型也包含以下特殊约束：

1. 绝对远程航向只分配给主基地航司或特定航司集团
2. 绝对远程航向必须分配宽体机
3. 特定航向不能分配宽体机
4. 航司只能分配给在现状中存在的航向（特例：国际离港的SAGPI航向可以分配给所有航司）

### 2.3 目标函数

离港航班分配模型使用与进港航班分配模型相同的多目标优化方法：

1. 最小化各航司集团未来波形与现状波形的差异
2. 最小化未来分布与当前分布的偏离度

```python
model.setObjective(airline_wave_deviation * 10000 + obj_expr, GRB.MINIMIZE)
```

## 3. 两个模型的联系

### 3.1 航班配对约束

进港航班分配模型和离港航班分配模型之间存在紧密的联系，主要体现在航班配对约束上：

1. **航司一致性**：离港航班的航司与相同ID的进港航班保持一致
   ```python
   # 在departure_assignment.py中
   for i, flight in day2_departures.iterrows():
       flight_id = flight['ID']
       matching_arrivals = arrival_assignments[arrival_assignments['ID'] == flight_id]
       if not matching_arrivals.empty:
           result_df.loc[i, '航司'] = matching_arrivals.iloc[0]['航司']
   ```

2. **离港航班配额约束**：进港航班分配模型考虑了对应离港航班的配额限制
   ```python
   # 在arrival_assignment.py中
   model.addConstr(dom_dep_expr + prev_dom_count <= dom_quota, f"dom_dep_quota_{airline}")
   model.addConstr(int_dep_expr + prev_int_count <= int_quota, f"int_dep_quota_{airline}")
   ```

3. **分配顺序**：系统先分配进港航班，再分配离港航班，确保一致性

### 3.2 共享约束类型

两个模型共享多种约束类型，包括：

1. **航司配额约束**：确保各航司分配的航班数量符合要求
2. **航向配额约束**：确保各航向分配的航班数量符合要求
3. **主航向约束**：保持各航司的主航向比例
4. **波形约束**：保持各小时的航班分布
5. **宽体机约束**：特定航向的宽体机分配规则
6. **绝对远程航向约束**：绝对远程航向只分配给特定航司

### 3.3 目标函数一致性

两个模型使用相同的多目标优化方法，包括：

1. 最小化各航司集团未来波形与现状波形的差异
2. 最小化未来分布与当前分布的偏离度

这确保了进港和离港航班分配的一致性和平衡性。

## 4. 数学模型总结

### 4.1 模型特点

1. **整数线性规划**：两个模型都是整数线性规划模型，使用二元决策变量
2. **多目标优化**：同时考虑多个目标，通过权重平衡各目标的重要性
3. **约束丰富**：包含多种约束条件，确保分配结果满足各种实际需求
4. **分阶段求解**：先分配进港航班，再分配离港航班，确保一致性

### 4.2 求解方法

两个模型都使用Gurobi求解器求解，并设置了MIP Gap参数控制求解精度：

```python
# 进港航班分配模型
if market_type == 'DOM':
    model.Params.MIPGap = DOM_ARR_GAP
else:
    model.Params.MIPGap = INT_ARR_GAP

# 离港航班分配模型
if market_type == 'DOM':
    model.Params.MIPGap = DOM_DEP_GAP
else:
    model.Params.MIPGap = INT_DEP_GAP
```

### 4.3 模型优势

1. **全局最优**：通过整数线性规划，可以找到全局最优解
2. **灵活性**：可以通过调整参数和约束条件，适应不同的需求
3. **可解释性**：模型结构清晰，约束条件明确，结果可解释
4. **一致性**：确保进港和离港航班分配的一致性和平衡性

### 4.4 潜在改进方向

1. **求解效率**：对于大规模问题，可以考虑使用启发式算法或分解方法提高求解效率
2. **约束松弛**：对于不可行问题，可以考虑松弛部分约束条件
3. **目标函数调整**：可以尝试不同的目标函数组合和权重设置
4. **模型集成**：考虑将进港和离港航班分配模型集成为一个统一的模型

通过这些数学优化模型，航司航向匹配优化系统能够为进港和离港航班分配最优的航司和航向组合，满足机场运营需求。