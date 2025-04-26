# config.py

# file config
CURRENT_STATUS_FILE = "2024_现状.xlsx"
ARR_HOURLY_STATS_FILE = "到达航班比例表.xlsx"
DEP_HOURLY_STATS_FILE = "出发航班比例表.xlsx"
PAIRING_FILE = "final_pairing_processed.xlsx"
OUTPUT_FILE = "final_result.xlsx"
FINAL_FILE = "final_result_processed.xlsx"

# MIPGap
DOM_ARR_GAP = 0.01
INT_ARR_GAP = 0.07

DOM_DEP_GAP = 0.02
INT_DEP_GAP = 0.01

# hourly wave bias
DOM_ARR_WAVE_BIAS = 0
INT_ARR_WAVE_BIAS = 0

DOM_DEP_WAVE_BIAS = 3
INT_DEP_WAVE_BIAS = 2

# wide body bias
DOM_ARR_WIDE_BIAS = 0
INT_ARR_WIDE_BIAS = 0

DOM_DEP_WIDE_BIAS = 2
INT_DEP_WIDE_BIAS = 2