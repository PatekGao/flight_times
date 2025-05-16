import os

import matplotlib.pyplot as plt
import pandas as pd

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 对于MacOS
plt.rcParams['axes.unicode_minus'] = False

def process_2024_data(file_path):
    # 读取2024年数据
    df = pd.read_excel(file_path)
    
    # 筛选flight type为PAX的数据
    df = df[df['flight type'] == 'PAX']
    
    # 从time列提取小时
    df['小时'] = pd.to_datetime(df['time']).dt.hour
    
    # 根据Direction判断到达/出发
    df['类型'] = df['Direction'].apply(lambda x: '到达' if x == 'Arrival' else '出发')
    
    # 根据flight type判断国内/国际
    df['市场'] = df['flightnature'].apply(lambda x: '国内' if x == 'DOM' else '国际')
    
    # 选择需要的列
    result = df[['小时', '市场', '类型', 'AirlineGroup']]
    
    # 创建完整的时间索引（0-23小时）
    hours = pd.DataFrame({'小时': range(24)})
    
    # 获取所有唯一的市场、类型和航司组合
    market_types = result[['市场', '类型', 'AirlineGroup']].drop_duplicates()
    
    # 创建完整的数据框
    complete_data = []
    for _, row in market_types.iterrows():
        market = row['市场']
        flight_type = row['类型']
        airline = row['AirlineGroup']
        
        # 获取该组合的现有数据
        group_data = result[(result['市场'] == market) & 
                          (result['类型'] == flight_type) & 
                          (result['AirlineGroup'] == airline)]
        
        # 统计每个小时的航班数
        hourly_counts = group_data.groupby('小时').size().reset_index(name='航班数')
        
        # 合并小时数据，确保所有小时都有记录
        merged_data = pd.merge(hours, hourly_counts, on='小时', how='left')
        merged_data['市场'] = market
        merged_data['类型'] = flight_type
        merged_data['AirlineGroup'] = airline
        merged_data['航班数'] = merged_data['航班数'].fillna(0)
        
        complete_data.append(merged_data)
    
    # 合并所有数据
    hourly_stats = pd.concat(complete_data, ignore_index=True)
    
    return hourly_stats

def process_2040_data(file_path):
    # 读取2040年数据
    df = pd.read_excel(file_path, sheet_name='配对航班')
    
    # 处理到达航班
    arr_data = df[['小时_ARR', '市场_ARR', '航司_ARR']].copy()
    arr_data.columns = ['小时', '市场', 'AirlineGroup']
    arr_data['类型'] = '到达'
    
    # 处理出发航班
    dep_data = df[['小时_DEP', '市场_DEP', '航司_DEP']].copy()
    dep_data.columns = ['小时', '市场', 'AirlineGroup']
    dep_data['类型'] = '出发'
    
    # 合并数据
    combined_data = pd.concat([arr_data, dep_data])
    
    # 将市场代码转换为中文
    market_map = {'INT': '国际', 'DOM': '国内'}
    combined_data['市场'] = combined_data['市场'].map(market_map)
    
    # 创建完整的时间索引（0-23小时）
    hours = pd.DataFrame({'小时': range(24)})
    
    # 获取所有唯一的市场、类型和航司组合
    market_types = combined_data[['市场', '类型', 'AirlineGroup']].drop_duplicates()
    
    # 创建完整的数据框
    complete_data = []
    for _, row in market_types.iterrows():
        market = row['市场']
        flight_type = row['类型']
        airline = row['AirlineGroup']
        
        # 获取该组合的现有数据
        group_data = combined_data[(combined_data['市场'] == market) & 
                                 (combined_data['类型'] == flight_type) & 
                                 (combined_data['AirlineGroup'] == airline)]
        
        # 统计每个小时的航班数
        hourly_counts = group_data.groupby('小时').size().reset_index(name='航班数')
        
        # 合并小时数据，确保所有小时都有记录
        merged_data = pd.merge(hours, hourly_counts, on='小时', how='left')
        merged_data['市场'] = market
        merged_data['类型'] = flight_type
        merged_data['AirlineGroup'] = airline
        merged_data['航班数'] = merged_data['航班数'].fillna(0)
        
        complete_data.append(merged_data)
    
    # 合并所有数据
    hourly_stats = pd.concat(complete_data, ignore_index=True)
    
    return hourly_stats

def plot_airline_waves(current_data, future_data, airline, output_dir):
    # 创建四个子图
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))
    fig.suptitle(f'{airline}航班波形对比（2024年vs2040年）', fontsize=16)
    
    # 设置颜色方案
    colors = {'2024年': 'blue', '2040年': 'red'}
    
    # 绘制国内到达航班波形
    current_domestic_arrivals = current_data[(current_data['市场'] == '国内') & 
                                           (current_data['类型'] == '到达') & 
                                           (current_data['AirlineGroup'] == airline)]
    future_domestic_arrivals = future_data[(future_data['市场'] == '国内') & 
                                         (future_data['类型'] == '到达') & 
                                         (future_data['AirlineGroup'] == airline)]
    
    if not current_domestic_arrivals.empty:
        ax1.plot(current_domestic_arrivals['小时'], current_domestic_arrivals['航班数'],
                label='2024年', color=colors['2024年'], linestyle='-', marker='o')
    if not future_domestic_arrivals.empty:
        ax1.plot(future_domestic_arrivals['小时'], future_domestic_arrivals['航班数'],
                label='2040年', color=colors['2040年'], linestyle='--', marker='s')
    
    ax1.set_title('国内到达航班波形')
    ax1.set_xlabel('小时')
    ax1.set_ylabel('航班数量')
    ax1.legend()
    ax1.grid(True)
    ax1.set_xticks(range(0, 24))
    
    # 绘制国内出发航班波形
    current_domestic_departures = current_data[(current_data['市场'] == '国内') & 
                                             (current_data['类型'] == '出发') & 
                                             (current_data['AirlineGroup'] == airline)]
    future_domestic_departures = future_data[(future_data['市场'] == '国内') & 
                                           (future_data['类型'] == '出发') & 
                                           (future_data['AirlineGroup'] == airline)]
    
    if not current_domestic_departures.empty:
        ax2.plot(current_domestic_departures['小时'], current_domestic_departures['航班数'],
                label='2024年', color=colors['2024年'], linestyle='-', marker='o')
    if not future_domestic_departures.empty:
        ax2.plot(future_domestic_departures['小时'], future_domestic_departures['航班数'],
                label='2040年', color=colors['2040年'], linestyle='--', marker='s')
    
    ax2.set_title('国内出发航班波形')
    ax2.set_xlabel('小时')
    ax2.set_ylabel('航班数量')
    ax2.legend()
    ax2.grid(True)
    ax2.set_xticks(range(0, 24))
    
    # 绘制国际到达航班波形
    current_international_arrivals = current_data[(current_data['市场'] == '国际') & 
                                                (current_data['类型'] == '到达') & 
                                                (current_data['AirlineGroup'] == airline)]
    future_international_arrivals = future_data[(future_data['市场'] == '国际') & 
                                              (future_data['类型'] == '到达') & 
                                              (future_data['AirlineGroup'] == airline)]
    
    if not current_international_arrivals.empty:
        ax3.plot(current_international_arrivals['小时'], current_international_arrivals['航班数'],
                label='2024年', color=colors['2024年'], linestyle='-', marker='o')
    if not future_international_arrivals.empty:
        ax3.plot(future_international_arrivals['小时'], future_international_arrivals['航班数'],
                label='2040年', color=colors['2040年'], linestyle='--', marker='s')
    
    ax3.set_title('国际到达航班波形')
    ax3.set_xlabel('小时')
    ax3.set_ylabel('航班数量')
    ax3.legend()
    ax3.grid(True)
    ax3.set_xticks(range(0, 24))
    
    # 绘制国际出发航班波形
    current_international_departures = current_data[(current_data['市场'] == '国际') & 
                                                  (current_data['类型'] == '出发') & 
                                                  (current_data['AirlineGroup'] == airline)]
    future_international_departures = future_data[(future_data['市场'] == '国际') & 
                                                (future_data['类型'] == '出发') & 
                                                (future_data['AirlineGroup'] == airline)]
    
    if not current_international_departures.empty:
        ax4.plot(current_international_departures['小时'], current_international_departures['航班数'],
                label='2024年', color=colors['2024年'], linestyle='-', marker='o')
    if not future_international_departures.empty:
        ax4.plot(future_international_departures['小时'], future_international_departures['航班数'],
                label='2040年', color=colors['2040年'], linestyle='--', marker='s')
    
    ax4.set_title('国际出发航班波形')
    ax4.set_xlabel('小时')
    ax4.set_ylabel('航班数量')
    ax4.legend()
    ax4.grid(True)
    ax4.set_xticks(range(0, 24))
    
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f'flight_wave_{airline}.png'), dpi=300, bbox_inches='tight')
    plt.close()

def reorganize_data_for_excel(current_data, future_data):
    # 创建小时索引（0-23）
    hours = pd.DataFrame({'小时': range(24)})
    
    # 获取所有航司
    all_airlines = sorted(set(current_data['AirlineGroup'].unique()) | set(future_data['AirlineGroup'].unique()))
    all_airlines = [airline for airline in all_airlines if airline != '货运']
    
    # 创建四个数据框：国内进港、国内离港、国际进港、国际离港
    sheets = {
        '国内进港': {'市场': '国内', '类型': '到达'},
        '国内离港': {'市场': '国内', '类型': '出发'},
        '国际进港': {'市场': '国际', '类型': '到达'},
        '国际离港': {'市场': '国际', '类型': '出发'}
    }
    
    result = {}
    
    for sheet_name, conditions in sheets.items():
        # 筛选当前sheet的数据
        current_sheet_data = current_data[
            (current_data['市场'] == conditions['市场']) & 
            (current_data['类型'] == conditions['类型'])
        ]
        future_sheet_data = future_data[
            (future_data['市场'] == conditions['市场']) & 
            (future_data['类型'] == conditions['类型'])
        ]
        
        # 创建结果数据框
        result_df = hours.copy()
        
        # 为每个航司添加列
        for airline in all_airlines:
            # 2024年数据
            current_airline_data = current_sheet_data[current_sheet_data['AirlineGroup'] == airline]
            if not current_airline_data.empty:
                result_df[f'{airline}_2024'] = result_df['小时'].map(
                    current_airline_data.set_index('小时')['航班数']
                ).fillna(0)
            
            # 2040年数据
            future_airline_data = future_sheet_data[future_sheet_data['AirlineGroup'] == airline]
            if not future_airline_data.empty:
                result_df[f'{airline}_2040'] = result_df['小时'].map(
                    future_airline_data.set_index('小时')['航班数']
                ).fillna(0)
        
        result[sheet_name] = result_df
    
    return result

def main():
    # 创建输出目录
    output_dir = 'flight_wave_plots_new'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 读取数据
    current_data = process_2024_data('2024_现状_WUH.xlsx')
    future_data = process_2040_data('final_result_processed_new.xlsx')
    
    # 重新组织数据并保存到Excel
    reorganized_data = reorganize_data_for_excel(current_data, future_data)
    
    with pd.ExcelWriter('flight_wave_analysis.xlsx') as writer:
        for sheet_name, data in reorganized_data.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # 获取所有航司集团（排除货运）
    all_airlines = sorted(set(current_data['AirlineGroup'].unique()) | set(future_data['AirlineGroup'].unique()))
    all_airlines = [airline for airline in all_airlines if airline != '货运']
    
    # 为每个航司集团生成波形图
    for airline in all_airlines:
        print(f"正在生成{airline}的波形图...")
        plot_airline_waves(current_data, future_data, airline, output_dir)
    
    print(f"所有波形图已生成完成！保存在 {output_dir} 文件夹中")
    print("处理后的数据已保存到 flight_wave_analysis.xlsx")

if __name__ == "__main__":
    main() 