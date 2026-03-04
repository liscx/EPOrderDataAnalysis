import pandas as pd
import yaml
import os
import sys


def load_config():
    """加载YAML配置文件"""
    try:
        with open("config.yaml", "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        print("错误：未找到 config.yaml 文件")
        sys.exit(1)


def run_core_metrics():
    print("\n[8/9] 正在提取全量核心概览指标...")

    # 1. 加载配置
    config = load_config()
    file_path = config['file_config']['output_file']

    # 严格对齐 YAML 时间
    conf_start = pd.to_datetime(config['analysis_period']['start_date'])
    conf_end = pd.to_datetime(config['analysis_period']['end_date'])

    target_month = conf_start.month
    month_label = f"{target_month}月"

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 2. 读取数据
    xls = pd.ExcelFile(file_path)
    all_sheets = {sn: xls.parse(sn) for sn in xls.sheet_names}

    if '清洗后数据' not in all_sheets:
        print("错误：Excel中缺少 '清洗后数据' 工作表")
        return

    df = all_sheets['清洗后数据'].copy()

    # 3. 预处理 (修复 Bug 的核心)
    # A. 核心：处理合并单元格产生的缺失 (向前填充日期和企业)
    df['订单日期'] = df['订单日期'].ffill()
    df['采购企业'] = df['采购企业'].ffill()
    df['订单号'] = df['订单号'].ffill()

    # B. 类型转换
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
    df['订单金额（元）'] = pd.to_numeric(df['订单金额（元）'], errors='coerce').fillna(0)
    df['数量'] = pd.to_numeric(df['数量'], errors='coerce').fillna(0)

    # C. 字符串清洗：去除商品名称前后的空格，防止 "苹果" 和 " 苹果" 被算作两类
    df['商品名称'] = df['商品名称'].astype(str).str.strip()

    # 4. 计算指标
    # --- A. 全量指标（整个表格的历史数据） ---
    total_purchasers = df['采购企业'].nunique()
    total_history_money = df['订单金额（元）'].sum()

    # 企业排行基础数据
    order_counts = df.groupby('采购企业')['订单号'].nunique()
    money_sums = df.groupby('采购企业')['订单金额（元）'].sum()

    max_orders_val, max_orders_ent = (order_counts.max(), order_counts.idxmax()) if not order_counts.empty else (0,
                                                                                                                 "无")
    max_money_val, max_money_ent = (money_sums.max(), money_sums.idxmax()) if not money_sums.empty else (0, "无")

    # --- B. 区间指标（指定日期范围内） ---
    mask = (df['订单日期'] >= conf_start) & (df['订单日期'] <= conf_end)
    df_period = df.loc[mask].copy()

    # 区间销售总额
    period_money_sum = df_period['订单金额（元）'].sum()
    # 区间商品销售件数
    total_items_count = int(df_period['数量'].sum())

    # 【修复 Bug】：区间商品种类
    # 去掉 '单价（元）' 的维度，因为同一商品在不同订单价格微调不应视为新种类
    # 使用 '商品名称' + '供应商' 作为去重逻辑最为稳妥
    unique_items_df = df_period.drop_duplicates(subset=['商品名称', '供应商'])
    unique_items_count = unique_items_df.shape[0]

    # 动态提取品类话术 (排除 nan 和 空值)
    valid_names = df_period[~df_period['商品名称'].isin(['nan', 'None', ''])]
    sample_cats = valid_names['商品名称'].str.split().str[0].unique()[:3]
    cat_str = "、".join(filter(None, sample_cats)) if len(sample_cats) > 0 else "通用商品"

    # 5. 构造结果表格
    metrics_data = [
        {"维度": "全局概览", "指标": "全量历史销售总额", "数值": round(total_history_money, 2), "单位": "元",
         "关联信息": "历史所有数据汇总"},
        {"维度": "全局概览", "指标": "采购企业总数", "数值": total_purchasers, "单位": "家", "关联信息": "---"},
        {"维度": "全局概览", "指标": "最高订单量企业", "数值": max_orders_val, "单位": "笔",
         "关联信息": max_orders_ent},
        {"维度": "全局概览", "指标": "最高采购金额企业", "数值": round(max_money_val, 2), "单位": "元",
         "关联信息": max_money_ent},

        {"维度": "区间概览", "指标": f"{month_label}销售总额", "数值": round(period_money_sum, 2), "单位": "元",
         "关联信息": f"{month_label}成交金额"},
        {"维度": "区间概览", "指标": f"{month_label}商品销售总数", "数值": total_items_count, "单位": "件",
         "关联信息": "订单商品数量累加"},
        {"维度": "区间概览", "指标": f"{month_label}交易商品种类", "数值": unique_items_count, "单位": "种",
         "关联信息": f"涵盖 {cat_str} 等品类"}
    ]
    metrics_df = pd.DataFrame(metrics_data)

    # 6. 控制台话术
    print("-" * 60)
    print(f"【报告话术参考】:")
    print(f"截止目前，全量历史销售总额累计达 {total_history_money:,.2f} 元。")
    print(
        f"其中 {month_label} 区间销售额为 {period_money_sum:,.2f} 元，共计售出商品 {total_items_count} 件，涵盖 {unique_items_count} 种商品。")
    print("-" * 60)

    # 7. 写回 Excel
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            metrics_df.to_excel(writer, sheet_name="核心概览指标", index=False)
        print(f"核心指标更新成功：已写入 Sheet [核心概览指标]")
    except Exception as e:
        # 如果 replace 模式失败（某些版本 openpyxl 不支持），回退到覆盖写入全表
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sn, s_df in all_sheets.items():
                if sn != "核心概览指标":
                    s_df.to_excel(writer, sheet_name=sn, index=False)
            metrics_df.to_excel(writer, sheet_name="核心概览指标", index=False)
        print(f"核心指标更新成功（全量覆盖模式）。")


if __name__ == "__main__":
    run_core_metrics()