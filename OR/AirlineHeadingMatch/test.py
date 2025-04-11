import pandas as pd
import numpy as np

def check_basic_info():
    """
    检查原始配对文件与分配结果文件中的基本信息是否一致
    比较的列包括：ID, 时间, 市场, 进出港, 机型, 日期
    """
    print("开始检查基本信息一致性...")

    # 文件路径
    original_file = "e:\\flight_times\\OR\\AirlineHeadingMatch\\final_pairing_processed.xlsx"
    result_file = "e:\\flight_times\\OR\\AirlineHeadingMatch\\4.6.1.45.xlsx"

    # 读取原始文件
    original_arrivals = pd.read_excel(original_file, sheet_name="到达航班")
    original_departures = pd.read_excel(original_file, sheet_name="出发航班")

    # 读取结果文件
    result_arrivals = pd.read_excel(result_file, sheet_name="到达航班")
    result_departures = pd.read_excel(result_file, sheet_name="出发航班")

    # 需要比较的列
    columns_to_check = ['ID', '时间', '市场', '进出港', '机型', '日期']

    # 检查到达航班
    print("\n检查到达航班基本信息...")
    check_dataframes(original_arrivals, result_arrivals, columns_to_check, "到达航班")

    # 检查出发航班
    print("\n检查出发航班基本信息...")
    check_dataframes(original_departures, result_departures, columns_to_check, "出发航班")
    
    # 检查航司和航向分配
    print("\n检查航司和航向分配情况...")
    check_assignments(result_arrivals, result_departures)

def check_dataframes(df1, df2, columns, sheet_name):
    """
    比较两个DataFrame中指定列的数据是否一致，考虑可能的行顺序变化
    
    Args:
        df1: 第一个DataFrame
        df2: 第二个DataFrame
        columns: 需要比较的列名列表
        sheet_name: 表名，用于输出信息
    """
    # 检查行数是否一致
    if df1.shape[0] != df2.shape[0]:
        print(f"警告：{sheet_name}行数不一致！原始文件: {df1.shape[0]}行, 结果文件: {df2.shape[0]}行")
        return
    
    # 创建用于比较的副本，并按ID排序
    if 'ID' in df1.columns and 'ID' in df2.columns:
        df1_sorted = df1.sort_values('ID').reset_index(drop=True)
        df2_sorted = df2.sort_values('ID').reset_index(drop=True)
        
        # 检查ID是否完全匹配
        if not df1_sorted['ID'].equals(df2_sorted['ID']):
            print(f"警告：{sheet_name}的ID列不完全匹配，可能有缺失或额外的航班")
            
            # 找出不匹配的ID
            missing_ids = set(df1_sorted['ID']) - set(df2_sorted['ID'])
            extra_ids = set(df2_sorted['ID']) - set(df1_sorted['ID'])
            
            if missing_ids:
                print(f"原始文件中存在但结果文件中缺失的ID: {missing_ids}")
            if extra_ids:
                print(f"结果文件中存在但原始文件中缺失的ID: {extra_ids}")
            
            # 只比较共有的ID
            common_ids = set(df1_sorted['ID']) & set(df2_sorted['ID'])
            df1_sorted = df1_sorted[df1_sorted['ID'].isin(common_ids)].sort_values('ID').reset_index(drop=True)
            df2_sorted = df2_sorted[df2_sorted['ID'].isin(common_ids)].sort_values('ID').reset_index(drop=True)
    else:
        print(f"警告：{sheet_name}中缺少ID列，无法确保正确比较")
        df1_sorted = df1.copy()
        df2_sorted = df2.copy()
    
    # 检查每一列是否存在
    for col in columns:
        if col not in df1_sorted.columns:
            print(f"警告：原始文件中不存在列 '{col}'")
            continue
        if col not in df2_sorted.columns:
            print(f"警告：结果文件中不存在列 '{col}'")
            continue
        
        # 比较列数据
        is_equal = df1_sorted[col].equals(df2_sorted[col])
        if not is_equal:
            # 找出不一致的行
            diff_mask = df1_sorted[col] != df2_sorted[col]
            diff_count = diff_mask.sum()
            
            print(f"列 '{col}' 数据不一致，共有 {diff_count} 行不同")
            
            # 显示前5个不一致的行
            if diff_count > 0:
                diff_indices = np.where(diff_mask)[0][:5]
                print(f"前{min(5, len(diff_indices))}个不一致的行:")
                for idx in diff_indices:
                    print(f"  ID {df1_sorted.loc[idx, 'ID']}: 原始值 = {df1_sorted.loc[idx, col]}, 结果值 = {df2_sorted.loc[idx, col]}")
        else:
            print(f"列 '{col}' 数据一致")

def check_assignments(arrivals_df, departures_df):
    """
    检查航司和航向的分配情况
    
    Args:
        arrivals_df: 到达航班DataFrame
        departures_df: 出发航班DataFrame
    """
    # 检查是否所有航班都分配了航司和航向
    missing_airline_arr = arrivals_df[arrivals_df['航司'].isna() | (arrivals_df['航司'] == "")].shape[0]
    missing_heading_arr = arrivals_df[arrivals_df['航向'].isna() | (arrivals_df['航向'] == "")].shape[0]
    
    missing_airline_dep = departures_df[departures_df['航司'].isna() | (departures_df['航司'] == "")].shape[0]
    missing_heading_dep = departures_df[departures_df['航向'].isna() | (departures_df['航向'] == "")].shape[0]
    
    print(f"到达航班中未分配航司的航班数: {missing_airline_arr}")
    print(f"到达航班中未分配航向的航班数: {missing_heading_arr}")
    print(f"出发航班中未分配航司的航班数: {missing_airline_dep}")
    print(f"出发航班中未分配航向的航班数: {missing_heading_dep}")
    
    # 创建ID到航司和日期的映射
    arr_airline_map = dict(zip(arrivals_df['ID'], arrivals_df['航司']))
    arr_date_map = dict(zip(arrivals_df['ID'], arrivals_df['日期']))
    
    # 检查每个离港航班的航司是否与相同ID的进港航班一致
    inconsistent_data = []
    
    for _, flight in departures_df.iterrows():
        flight_id = flight['ID']
        dep_airline = flight['航司']
        dep_date = flight['日期']
        
        if flight_id in arr_airline_map:
            arr_airline = arr_airline_map[flight_id]
            arr_date = arr_date_map[flight_id]
            if dep_airline != arr_airline and dep_airline != "" and arr_airline != "":
                inconsistent_data.append((flight_id, arr_airline, dep_airline, arr_date, dep_date))
    
    # 按ID排序不一致的示例
    inconsistent_data.sort(key=lambda x: x[0])
    
    print(f"\n所有进出港航班航司不一致的数量: {len(inconsistent_data)}")
    if inconsistent_data:
        print("不一致示例(按ID排序):")
        for flight_id, arr_airline, dep_airline, arr_date, dep_date in inconsistent_data:
            print(f"  ID {flight_id}: 进港航司 = {arr_airline} (日期 = {arr_date}), 出港航司 = {dep_airline} (日期 = {dep_date})")
    
    # 按日期分组统计不一致数量
    date_pairs = {}
    for flight_id, arr_airline, dep_airline, arr_date, dep_date in inconsistent_data:
        date_pair = (arr_date, dep_date)
        if date_pair not in date_pairs:
            date_pairs[date_pair] = 0
        date_pairs[date_pair] += 1
    
    print("\n按日期对分组的不一致数量:")
    for (arr_date, dep_date), count in sorted(date_pairs.items()):
        print(f"  进港日期 = {arr_date}, 出港日期 = {dep_date}: {count}个不一致")
    
    # 统计各航司和航向的分配情况
    print("\n到达航班航司分配统计:")
    print(arrivals_df['航司'].value_counts().head(10))
    
    print("\n到达航班航向分配统计:")
    print(arrivals_df['航向'].value_counts().head(10))
    
    print("\n出发航班航司分配统计:")
    print(departures_df['航司'].value_counts().head(10))
    
    print("\n出发航班航向分配统计:")
    print(departures_df['航向'].value_counts().head(10))
    
    # 按日期分组统计航司和航向分配
    print("\n按日期分组的航司分配统计:")
    for date in sorted(arrivals_df['日期'].unique()):
        date_df = arrivals_df[arrivals_df['日期'] == date]
        print(f"  到达航班日期 = {date}:")
        print(f"    航司分配: {date_df['航司'].value_counts().head(5).to_dict()}")
        print(f"    航向分配: {date_df['航向'].value_counts().head(5).to_dict()}")
    
    for date in sorted(departures_df['日期'].unique()):
        date_df = departures_df[departures_df['日期'] == date]
        print(f"  出发航班日期 = {date}:")
        print(f"    航司分配: {date_df['航司'].value_counts().head(5).to_dict()}")
        print(f"    航向分配: {date_df['航向'].value_counts().head(5).to_dict()}")

def main():
    print("开始执行测试...")
    check_basic_info()
    print("\n测试完成！")

if __name__ == "__main__":
    main()