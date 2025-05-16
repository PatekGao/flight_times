import pandas as pd
from OR.TableGeneration.pre_task2 import peak_results
from OR.TableGeneration.utils import DOM_MAX, ARR_DOM_MAX, DEP_DOM_MAX, INT_MAX, ARR_INT_MAX, DEP_INT_MAX, xls

df2 = pd.read_excel(xls, sheet_name="其它参数设置")

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
