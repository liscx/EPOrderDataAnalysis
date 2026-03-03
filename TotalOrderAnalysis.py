import pandas as pd
import yaml
import os


def load_config():
    """加载YAML配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_stats_df(df_source, mode):
    """核心统计逻辑：构建全量骨架，确保0订单输出，并空行分隔"""
    if df_source.empty:
        return pd.DataFrame()

    # 1. 数据清洗与预处理
    df = df_source.copy()
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
    df['订单号'] = df['订单号'].ffill()
    df['专区名称'] = df['专区名称'].ffill()
    df['月份'] = df['订单日期'].dt.strftime('%Y年%m月')

    # 获取全量维度集合
    all_months = sorted(df['月份'].dropna().unique())
    all_zones = sorted(df['专区名称'].dropna().unique())

    # 2. 核心统计：先聚合实际有的数据
    actual_stats = df.groupby(['月份', '专区名称']).agg(
        订单数量=('订单号', 'nunique'),
        交易金额=('订单金额（元）', 'sum')
    ).reset_index()

    final_rows = []

    if mode == "时间维度":
        # 严格遍历：每个月都要有所有专区
        for m in all_months:
            m_total_qty = 0
            m_total_money = 0

            for z in all_zones:
                # 在聚合结果中匹配
                match = actual_stats[(actual_stats['月份'] == m) & (actual_stats['专区名称'] == z)]

                # 即使匹配不到，也强制输出 0
                qty = int(match['订单数量'].iloc[0]) if not match.empty else 0
                money = float(match['交易金额'].iloc[0]) if not match.empty else 0.0

                final_rows.append({'时间/专区': m, '明细项': z, '订单数量': qty, '交易金额(元)': money})
                m_total_qty += qty
                m_total_money += money

            # 插入该月小计
            final_rows.append(
                {'时间/专区': f"{m} 小计", '明细项': '---', '订单数量': m_total_qty, '交易金额(元)': m_total_money})
            # 插入空行（阅读性优化）
            final_rows.append({'时间/专区': '', '明细项': '', '订单数量': '', '交易金额(元)': ''})

    else:  # 专区维度
        # 严格遍历：每个专区都要有所有月份
        for z in all_zones:
            z_total_qty = 0
            z_total_money = 0

            for m in all_months:
                match = actual_stats[(actual_stats['月份'] == m) & (actual_stats['专区名称'] == z)]

                qty = int(match['订单数量'].iloc[0]) if not match.empty else 0
                money = float(match['交易金额'].iloc[0]) if not match.empty else 0.0

                final_rows.append({'专区/时间': z, '明细项': m, '订单数量': qty, '交易金额(元)': money})
                z_total_qty += qty
                z_total_money += money

            # 插入该专区小计
            final_rows.append(
                {'专区/时间': f"{z} 小计", '明细项': '---', '订单数量': z_total_qty, '交易金额(元)': z_total_money})
            # 插入空行（阅读性优化）
            final_rows.append({'专区/时间': '', '明细项': '', '订单数量': '', '交易金额(元)': ''})

    return pd.DataFrame(final_rows)


def run_analysis():
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 1. 加载数据源
    xls = pd.ExcelFile(file_path)
    all_sheets = {sn: xls.parse(sn) for sn in xls.sheet_names}

    df_clean = all_sheets.get('清洗后数据', pd.DataFrame())
    df_period = all_sheets.get('指定区间数据', pd.DataFrame())

    # 2. 定义任务
    # 特别注意：全量报表的“骨架”应基于全量数据的维度，区间报表的“骨架”基于区间数据的维度
    tasks = [
        {"df": df_clean, "mode": "时间维度", "name": "汇总_时间_全量"},
        {"df": df_period, "mode": "时间维度", "name": "汇总_时间_区间"},
        {"df": df_clean, "mode": "专区维度", "name": "汇总_专区_全量"},
        {"df": df_period, "mode": "专区维度", "name": "汇总_专区_区间"}
    ]

    analysis_results = {}
    for task in tasks:
        if not task["df"].empty:
            print(f"处理中: {task['name']} (强制补零 + 空行分离)...")
            analysis_results[task["name"]] = get_stats_df(task["df"], task["mode"])

    # 3. 写回 Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for sn, s_df in all_sheets.items():
            if sn not in analysis_results.keys():
                s_df.to_excel(writer, sheet_name=sn, index=False)

        for sn, r_df in analysis_results.items():
            r_df.to_excel(writer, sheet_name=sn, index=False)

    print(f"分析完成！四张汇总表已按严格格式输出至 '{file_path}'。")


if __name__ == "__main__":
    run_analysis()