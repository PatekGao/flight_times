import pandas as pd
from airline_heading_match import get_hour_from_time

def process_excel(INPUT_FILE, OUTPUT_FILE):
    # 读取Excel文件
    arrival_flights = pd.read_excel(INPUT_FILE, sheet_name="到达航班")
    departure_flights = pd.read_excel(INPUT_FILE, sheet_name="出发航班")
    
    # 补齐小时列
    arrival_flights.loc[:, '小时'] = arrival_flights['时间'].apply(get_hour_from_time)
    departure_flights.loc[:, '小时'] = departure_flights['时间'].apply(get_hour_from_time)
    
    # 按ID合并到达和出发航班
    merged_flights = pd.merge(
        arrival_flights, 
        departure_flights, 
        on='ID', 
        how='outer',
        suffixes=('_ARR', '_DEP')
    )
    
    # 将结果写入新的Excel文件
    with pd.ExcelWriter(OUTPUT_FILE) as writer:
        arrival_flights.to_excel(writer, sheet_name="到达航班", index=False)
        departure_flights.to_excel(writer, sheet_name="出发航班", index=False)
        merged_flights.to_excel(writer, sheet_name="配对航班", index=False)
    
    print(f"处理完成！结果已保存到 {OUTPUT_FILE}")

