from arrival_assignment import *
from departure_assignment import *
from process_final_result import *
from utils import load_data, complete_day1_assignments, complete_day3_assignments
from config import OUTPUT_FILE, FINAL_FILE


def main():
    # 加载数据
    current_status, arrival_flights, departure_flights, hourly_dom_stats, hourly_int_stats, hourly_dom_dep_stats, hourly_int_dep_stats = load_data()

    # 分配国内进港航班，传入离港航班数据
    dom_arr_assignments, dom_dep_counts, int_dep_counts = assign_arrival_flights(
        current_status, arrival_flights, 'DOM', hourly_dom_stats, hourly_int_stats, departure_flights
    )

    # 分配国际进港航班，传入离港航班数据和之前的分配结果
    int_arr_assignments, _, _ = assign_arrival_flights(
        current_status, arrival_flights, 'INT', hourly_dom_stats, hourly_int_stats,
        departure_flights, dom_dep_counts, int_dep_counts
    )

    # 合并进港航班分配结果
    all_day2_arr_assignments = pd.concat([dom_arr_assignments, int_arr_assignments])

    # 分配国内离港航班
    dom_dep_assignments = assign_departure_flights(all_day2_arr_assignments, departure_flights, current_status, 'DOM',
                                                   hourly_dom_dep_stats, hourly_int_dep_stats)

    # 分配国际离港航班
    int_dep_assignments = assign_departure_flights(all_day2_arr_assignments, departure_flights, current_status, 'INT',
                                                   hourly_dom_dep_stats, hourly_int_dep_stats)

    # 合并离港航班分配结果
    all_day2_dep_assignments = pd.concat([dom_dep_assignments, int_dep_assignments])

    # 补齐日期为1的进港航班分配
    day1_arr_assignments = complete_day1_assignments(all_day2_dep_assignments,
                                                     arrival_flights, departure_flights)

    # 补齐日期为3的离港航班分配
    day3_dep_assignments = complete_day3_assignments(all_day2_arr_assignments, departure_flights)

    # 合并所有分配结果
    all_arr_assignments = pd.concat([day1_arr_assignments, all_day2_arr_assignments])
    all_dep_assignments = pd.concat([all_day2_dep_assignments, day3_dep_assignments])

    # 将结果写回Excel文件
    with pd.ExcelWriter(OUTPUT_FILE) as writer:
        all_arr_assignments.to_excel(writer, sheet_name="到达航班", index=False)
        all_dep_assignments.to_excel(writer, sheet_name="出发航班", index=False)

    print(f"分配完成！结果已保存到 {OUTPUT_FILE}")

    process_excel(OUTPUT_FILE, FINAL_FILE)


if __name__ == "__main__":
    main()
