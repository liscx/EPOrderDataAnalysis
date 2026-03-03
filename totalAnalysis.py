# step2 全局汇总
import pandas as pd
import os

# ================= 配置区域 =================
START_DATE = '2025-06-01'
END_DATE = '2026-02-28'
INPUT_FILE = '清洗结果_汇总.xlsx'
OUTPUT_FILE = '清洗结果_汇总.xlsx'

# 核心配置项：可选值为 "时间维度" 或 "专区维度"
REPORT_MODE = "专区维度"


# ===========================================

def run_analysis():
    if not os.path.exists(INPUT_FILE):
        print(f"错误：找不到文件 {INPUT_FILE}")
        return

    # 1. 加载数据
    with pd.ExcelFile(INPUT_FILE) as xls:
        df = pd.read_excel(xls, sheet_name='清洗后数据')
        df_removed = pd.read_excel(xls, sheet_name='被清洗的测试数据')

    # 2. 日期处理与筛选
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
    start_dt = pd.to_datetime(START_DATE)
    end_dt = pd.to_datetime(END_DATE)
    df_period = df[(df['订单日期'] >= start_dt) & (df['订单日期'] <= end_dt)].copy()

    if df_period.empty:
        print(f"警告：{START_DATE} 至 {END_DATE} 期间无数据")
        return

    # 3. 准备统计维度
    all_months_dt = pd.date_range(start=start_dt, end=end_dt, freq='MS')
    all_zones = sorted(df_period['专区名称'].dropna().unique().tolist())

    # 4. 总体概况输出
    total_orders = df_period['订单号'].nunique()
    total_money = df_period['订单金额（元）'].sum()
    print(f"--- 统计报告 ({START_DATE} 至 {END_DATE}) ---")
    print(f"当前模式：【{REPORT_MODE}】")
    print(f"全站累计总订单: {total_orders} 笔，全站累计总金额: {total_money:,.2f} 元")
    print("-" * 65)

    # 5. 执行逻辑
    if REPORT_MODE == "时间维度":
        for dt in all_months_dt:
            m_label = dt.strftime('%Y年%m月')
            print(f"\n【{m_label}】交易明细：")

            df_m = df_period[(df_period['订单日期'].dt.year == dt.year) & (df_period['订单日期'].dt.month == dt.month)]
            zone_stats = df_m.groupby('专区名称').agg(count=('订单号', 'nunique'), money=('订单金额（元）', 'sum'))

            # 记录月度小计
            m_subtotal_orders = df_m['订单号'].nunique()
            m_subtotal_money = df_m['订单金额（元）'].sum()

            for i, z_name in enumerate(all_zones, start=1):
                cnt = int(zone_stats.loc[z_name, 'count']) if z_name in zone_stats.index else 0
                amt = zone_stats.loc[z_name, 'money'] if z_name in zone_stats.index else 0.0
                print(f"  {i}. {z_name}：{cnt} 笔订单，{amt:,.2f} 元")

            # 输出月度汇总
            print(f"  >> {m_label}汇总：总计 {m_subtotal_orders} 个订单，累计金额 {m_subtotal_money:,.2f} 元")

    elif REPORT_MODE == "专区维度":
        zone_rank = df_period.groupby('专区名称')['订单金额（元）'].sum().sort_values(ascending=False).index

        for i, z_name in enumerate(zone_rank, start=1):
            df_z = df_period[df_period['专区名称'] == z_name]
            z_total_orders = df_z['订单号'].nunique()
            z_total_money = df_z['订单金额（元）'].sum()

            print(f"\n{i}. {z_name} 明细：")

            month_stats = df_z.groupby(df_z['订单日期'].dt.strftime('%Y年%m月')).agg(count=('订单号', 'nunique'),
                                                                                     money=('订单金额（元）', 'sum'))

            for dt in all_months_dt:
                m_label = dt.strftime('%Y年%m月')
                cnt = int(month_stats.loc[m_label, 'count']) if m_label in month_stats.index else 0
                amt = month_stats.loc[m_label, 'money'] if m_label in month_stats.index else 0.0
                print(f"   * {m_label}：{cnt} 笔订单，{amt:,.2f} 元")

            # 输出专区汇总
            print(f"   >> {z_name}总计：总计 {z_total_orders} 个订单，累计金额 {z_total_money:,.2f} 元")

    # 6. 保存数据
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='清洗后数据', index=False)
        df_removed.to_excel(writer, sheet_name='被清洗的测试数据', index=False)
        df_period.to_excel(writer, sheet_name='指定区间数据', index=False)


if __name__ == "__main__":
    run_analysis()