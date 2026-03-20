import pandas as pd
import yaml
import os


def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_enhanced_metrics():
    print("正在基于 [清洗后数据] 修复收货订单统计逻辑并提取话术指标...")
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 1. 加载数据
    xls = pd.ExcelFile(file_path)
    df_raw = xls.parse('清洗后数据')
    df_time_range = xls.parse('汇总_时间_区间')
    df_time_all = xls.parse('汇总_时间_全量')
    df_zone_all = xls.parse('汇总_专区_全量')
    df_buyer_sum = xls.parse('采购企业汇总表')
    df_sup_sum = xls.parse('供应商汇总表')

    # --- 基础处理：填充合并单元格导致的空值 ---
    fill_cols = ['订单号', '订单日期', '订单状态', '采购企业', '供应商', '订单金额（元）']
    for col in fill_cols:
        if col in df_raw.columns:
            df_raw[col] = df_raw[col].ffill()

    # 统一格式
    df_raw['订单日期'] = pd.to_datetime(df_raw['订单日期'], errors='coerce')
    df_raw['订单金额（元）'] = pd.to_numeric(df_raw['订单金额（元）'], errors='coerce').fillna(0)

    # 获取配置参数
    conf_start = pd.to_datetime(config['analysis_period']['start_date'])
    target_month = conf_start.month
    target_year = conf_start.year
    last_month_str = (conf_start - pd.offsets.MonthBegin(1)).strftime('%Y年%m月')

    m_list = []

    # --- 1. 概况 (话术1) - 重点修复去重逻辑 ---
    # A. 全量订单统计
    df_all_orders = df_raw.drop_duplicates(subset=['订单号'])
    total_buyer = df_raw['采购企业'].nunique()
    total_sup = df_raw['供应商'].nunique()

    ec_keys = ['得力', '齐心', '苏宁', '史泰博', '欧菲斯', '紫迈', '鑫方盛', '震坤行']
    is_ec = df_raw['供应商'].apply(lambda x: any(k in str(x) for k in ec_keys))
    ec_count = df_raw[is_ec]['供应商'].nunique()

    all_order_count = df_all_orders['订单号'].nunique()
    all_order_money = df_all_orders['订单金额（元）'].sum()

    # B. 完成收货订单统计（修复点：先筛选状态，再订单号去重，最后求和金额）
    # 注意：此处你可以根据实际状态名调整，如 ['收货完成']
    target_status = ['收货完成']
    df_done_rows = df_raw[df_raw['订单状态'].isin(target_status)]
    df_done_unique = df_done_rows.drop_duplicates(subset=['订单号'])

    done_count = df_done_unique['订单号'].nunique()
    done_money = df_done_unique['订单金额（元）'].sum()

    m_list.append(["话术1", "概况", f"{total_buyer}|{total_sup}|{ec_count}|{total_sup - ec_count}",
                   f"{all_order_count}|{all_order_money:.2f}|{done_count}|{done_money:.2f}"])

    # --- 2. 累计专区 (话术2) ---
    df_zone_summary = df_zone_all[df_zone_all['专区/时间'].str.contains('小计', na=False)].copy()
    total_orders_all = df_zone_summary['订单数量'].sum()
    active_zones_count = len(df_zone_summary[df_zone_summary['订单数量'] > 0])
    target_zones = ["中国煤地电子商城", "新疆阳光采购平台", "大连市阳光采购服务平台", "邯郸市阳光优采平台"]
    zone_results = []
    for zone_name in target_zones:
        match = df_zone_summary[df_zone_summary['专区/时间'].str.contains(zone_name, na=False)]
        if not match.empty:
            count = match['订单数量'].values[0]
            percent = (count / total_orders_all * 100) if total_orders_all > 0 else 0
            zone_results.append(f"{int(count)}|{percent:.2f}%")
        else:
            zone_results.append("0|0.00%")
    m_list.append(["话术2", "累计专区", active_zones_count, f"{int(total_orders_all)}|" + "|".join(zone_results)])

    # --- 3. 本月表现 (话术3) ---
    # 获取区间汇总小计行
    tm_subtotal_mask = df_time_range['时间/专区'].str.contains('小计', na=False)
    if not df_time_range[tm_subtotal_mask].empty:
        tm_row = df_time_range[tm_subtotal_mask].iloc[0]
        tm_money = tm_row['交易金额(元)']
        tm_count = tm_row['订单数量']
    else:
        tm_money, tm_count = 0, 0

    lm_data = df_time_all[df_time_all['时间/专区'].str.contains(f"{last_month_str} 小计", na=False)]
    lm_money = lm_data['交易金额(元)'].values[0] if not lm_data.empty else 0
    zone_rank = df_time_range[~df_time_range['明细项'].str.contains('---', na=False)].sort_values('交易金额(元)',
                                                                                                  ascending=False)

    def get_top_zone_info(rank_df, pos):
        if len(rank_df) > pos:
            row = rank_df.iloc[pos]
            return f"{row['明细项']}|{int(row['订单数量'])}|{row['交易金额(元)']:.2f}"
        return "无|0|0.00"

    m_list.append(["话术3", "本月表现", f"{target_month}|{len(zone_rank)}|{int(tm_count)}",
                   f"{tm_money:.2f}|{lm_money:.2f}|{get_top_zone_info(zone_rank, 0)}|{get_top_zone_info(zone_rank, 1)}"])

    # --- 4. 历史最值 (话术4) ---
    if not df_buyer_sum.empty:
        max_o_row = df_buyer_sum.sort_values('订单数量', ascending=False).iloc[0]
        max_m_row = df_buyer_sum.sort_values('订单总额（元）', ascending=False).iloc[0]
        m_list.append(["话术4", "最值", total_buyer,
                       f"{int(max_o_row['订单数量'])}|{max_o_row['采购企业']}|{max_o_row['订单总额（元）']:.2f}|{max_o_row['专区名称']}|"
                       f"{max_m_row['订单总额（元）']:.2f}|{max_m_row['采购企业']}|{int(max_m_row['订单数量'])}|{max_m_row['专区名称']}"])

    # --- 5 & 8 占位 ---
    m_list.append(["话术5", "新采购人", "待计算", "建议从首次订单日期判断"])
    m_list.append(["话术8", "新供应商", "待计算", "建议从首次订单日期判断"])

    # --- 6. 交易供应商分类 (话术6) ---
    m_list.append(["话术6", "交易供应商", total_sup, f"{ec_count}|{total_sup - ec_count}"])

    # --- 7. 本地表现 (话术7) ---
    df_local = df_sup_sum[~df_sup_sum['供应商'].apply(lambda x: any(k in str(x) for k in ec_keys))].copy()
    df_local = df_local.sort_values('订单总额（元）', ascending=False)
    local_total_amt = df_local['订单总额（元）'].sum()
    m_list.append(["话术7", "本地表现", len(df_local),
                   f"{local_total_amt:.2f}|{get_top_zone_info(df_local.rename(columns={'供应商': '明细项', '订单总额（元）': '交易金额(元)'}), 0)}"])

    # --- 9. 商品概括 (话术9) ---
    this_month_mask = (df_raw['订单日期'].dt.month == target_month) & (df_raw['订单日期'].dt.year == target_year)
    df_this_month = df_raw[this_month_mask].copy()
    if not df_this_month.empty:
        valid_items = df_this_month[~df_this_month['商品名称'].astype(str).isin(['nan', '', 'None', '汇总', '合计'])]
        unique_item_count = valid_items['商品名称'].nunique()
        sample_items = "、".join(valid_items['商品名称'].astype(str).unique()[:5])
    else:
        unique_item_count = 0
        sample_items = "无"
    m_list.append(["话术9", "商品概括", unique_item_count, sample_items])

    # --- 保存 ---
    metrics_df = pd.DataFrame(m_list, columns=['维度', '指标', '数值', '关联信息'])
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            metrics_df.to_excel(writer, sheet_name="核心话术数据", index=False)
        print(">>> 核心话术指标提取完成！收货重复统计已修复。")
    except Exception as e:
        print(f"写入失败: {e}")


if __name__ == "__main__":
    run_enhanced_metrics()