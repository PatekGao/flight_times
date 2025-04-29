import pandas as pd

from OR.TableGeneration.pre_task2 import peak_results


# 预处理函数：将时间字符串转换为分钟数
def time_to_minutes(t_str):
    h, m = map(int, t_str.split(':'))
    return h * 60 + m  # 机型和市场配额约束（日期均为2）


xls = pd.ExcelFile("python时刻表输入v3.xlsx")

# 读取指定的 sheet
df1 = pd.read_excel(xls, sheet_name="规划静态时刻表")
df2 = pd.read_excel(xls, sheet_name="其它参数设置")
df3 = pd.read_excel(xls, sheet_name="现状动态统计")
df4 = pd.read_excel(xls, sheet_name="现状时刻表")

print("Reading excel data...")
ref_arr_dom = df3['DOM ARR'].values
ref_arr_int = df3['INT ARR'].values
ref_dep_dom = df3['DOM DEP'].values
ref_dep_int = df3['INT DEP'].values

H_arr_dom = df1["ARR DOM"][:24].values
H_arr_int = df1["ARR INT"][:24].values
H_dep_dom = df1["DEP DOM"][:24].values
H_dep_int = df1["DEP INT"][:24].values

ARR_DOM_MAX = df1["ARR DOM"][26]
ARR_INT_MAX = df1["ARR INT"][26]
DEP_DOM_MAX = df1["DEP DOM"][26]
DEP_INT_MAX = df1["DEP INT"][26]

ARR_MAX = df1["ARR"][26]
DEP_MAX = df1["DEP"][26]
DOM_MAX = df1["DOM"][26]
INT_MAX = df1["INT"][26]
TOT_MAX = df1["TOT"][26]

ARR_LIMIT = df1["ARR"][27]
DEP_LIMIT = df1["DEP"][27]
TOT_LIMIT = df1["TOT"][27]

ARR_15_LIMIT = df1["ARR"][28]
DEP_15_LIMIT = df1["DEP"][28]
TOT_15_LIMIT = df1["TOT"][28]

DOM_MIX_p = df2.iloc[20, 3]
INT_MIX_p = df2.iloc[20, 4]

DOM_INT_MIX_p = df2.iloc[24, 3]

quotas = {
    ('DOM', 'B'): {'ARR': df2.iloc[2, 3], 'DEP': df2.iloc[2, 5]},
    ('DOM', 'C'): {'ARR': df2.iloc[3, 3], 'DEP': df2.iloc[3, 5]},
    ('DOM', 'D'): {'ARR': df2.iloc[4, 3], 'DEP': df2.iloc[4, 5]},
    ('DOM', 'E'): {'ARR': df2.iloc[5, 3], 'DEP': df2.iloc[5, 5]},
    ('DOM', 'F'): {'ARR': df2.iloc[6, 3], 'DEP': df2.iloc[6, 5]},
    ('INT', 'B'): {'ARR': df2.iloc[2, 4], 'DEP': df2.iloc[2, 6]},
    ('INT', 'C'): {'ARR': df2.iloc[3, 4], 'DEP': df2.iloc[3, 6]},
    ('INT', 'D'): {'ARR': df2.iloc[4, 4], 'DEP': df2.iloc[4, 6]},
    ('INT', 'E'): {'ARR': df2.iloc[5, 4], 'DEP': df2.iloc[5, 6]},
    ('INT', 'F'): {'ARR': df2.iloc[6, 4], 'DEP': df2.iloc[6, 6]},
}
quotas = {k: v for k, v in quotas.items() if not (pd.isna(v['ARR']) and pd.isna(v['DEP']))}

DOM_window, DOM_count = peak_results['国内双向']
DOM_ARR_window, DOM_ARR_count = peak_results['国内进港']
DOM_DEP_window, DOM_DEP_count = peak_results['国内离港']
INT_window, INT_count = peak_results['国际双向']
INT_ARR_window, INT_ARR_count = peak_results['国际进港']
INT_DEP_window, INT_DEP_count = peak_results['国际离港']
peak_configs = [
    # 国内双向
    {
        'name': '国内双向',
        'market': 'DOM',
        'direction': 'both',
        'start_time': DOM_window[:5],
        'end_time': DOM_window[6:],
        'total': DOM_MAX,
        'ratios': {'C': df2.iloc[12, 5], 'E': df2.iloc[14, 5], 'F': df2.iloc[15, 5]}
    },
    # 国内进港
    {
        'name': '国内进港',
        'market': 'DOM',
        'direction': 'ARR',
        'start_time': DOM_ARR_window[:5],
        'end_time': DOM_ARR_window[6:],
        'total': ARR_DOM_MAX,
        'ratios': {'C': df2.iloc[12, 3], 'E': df2.iloc[14, 3], 'F': df2.iloc[15, 3]}
    },
    # 国内离港
    {
        'name': '国内离港',
        'market': 'DOM',
        'direction': 'DEP',
        'start_time': DOM_DEP_window[:5],
        'end_time': DOM_DEP_window[6:],
        'total': DEP_DOM_MAX,
        'ratios': {'C': df2.iloc[12, 4], 'E': df2.iloc[14, 4], 'F': df2.iloc[15, 4]}
    },
    # 国际双向
    {
        'name': '国际双向',
        'market': 'INT',
        'direction': 'both',
        'start_time': INT_window[:5],
        'end_time': INT_window[6:],
        'total': INT_MAX,
        'ratios': {'C': df2.iloc[12, 8], 'E': df2.iloc[14, 8], 'F': df2.iloc[15, 8]}
    },
    # 国际进港
    {
        'name': '国际进港',
        'market': 'INT',
        'direction': 'ARR',
        'start_time': INT_ARR_window[:5],
        'end_time': INT_ARR_window[6:],
        'total': ARR_INT_MAX,
        'ratios': {'C': df2.iloc[12, 6], 'E': df2.iloc[14, 6], 'F': df2.iloc[15, 6]}
    },
    # 国际离港
    {
        'name': '国际离港',
        'market': 'INT',
        'direction': 'DEP',
        'start_time': INT_DEP_window[:5],
        'end_time': INT_DEP_window[6:],
        'total': DEP_INT_MAX,
        'ratios': {'C': df2.iloc[12, 7], 'E': df2.iloc[14, 7], 'F': df2.iloc[15, 7]}
    }
]

min_times = {
    ('DOM', 'C'): df2.iloc[29, 3],
    ('DOM', 'E'): df2.iloc[31, 3],
    ('DOM', 'F'): df2.iloc[32, 3],
    ('INT', 'C'): df2.iloc[29, 4],
    ('INT', 'E'): df2.iloc[31, 4],
    ('INT', 'F'): df2.iloc[32, 4],
}

dep_hour_distribution = [(row[0], row[1], row[2]) for row in df2.iloc[37:61, 2:5].values]
print("Finished reading! Now out of utils.")
if __name__ == "__main__":
    # test
    print(ref_arr_dom)
    print(ref_arr_int)
    print(ref_dep_dom)
    print(ref_dep_int)
    print(len(ref_arr_dom))

    print(H_arr_dom)
    print(H_arr_int)
    print(H_dep_dom)
    print(H_dep_int)
    print(H_dep_dom[0])

    print(ARR_DOM_MAX)
    print(ARR_INT_MAX)
    print(DEP_DOM_MAX)
    print(DEP_INT_MAX)

    print(ARR_MAX)
    print(DEP_MAX)
    print(DOM_MAX)
    print(INT_MAX)
    print(TOT_MAX)

    print(ARR_LIMIT)
    print(DEP_LIMIT)
    print(TOT_LIMIT)

    print(ARR_15_LIMIT)
    print(DEP_15_LIMIT)
    print(TOT_15_LIMIT)

    print(quotas)

    print(DOM_MIX_p)
    print(INT_MIX_p)

    print(DOM_INT_MIX_p)

    print(peak_configs)

    print(min_times)

    print(dep_hour_distribution)