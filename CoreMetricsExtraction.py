import pandas as pd
import yaml
import os


def load_config():
    """加载YAML配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_core_metrics():
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
    df = all_sheets['清洗后数据'].copy()

    # 3. 预处理
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
    df['订单号'] = df['订单号'].ffill()
    df['采购企业'] = df['采购企业'].ffill()
    df['供应商'] = df['供应商'].ffill()
    df['商品名称'] = df['商品名称'].astype(str)

    # 4. 计算指标
    # A. 全量指标
    total_purchasers = df['采购企业'].nunique()
    order_counts = df.groupby('采购企业')['订单号'].nunique()
    money_sums = df.groupby('采购企业')['订单金额（元）'].sum()
    max_orders_val, max_orders_ent = order_counts.max(), order_counts.idxmax()
    max_money_val, max_money_ent = money_sums.max(), money_sums.idxmax()

    # B. 【关键修正】区间统计：同样的商品只计一次
    mask = (df['订单日期'] >= conf_start) & (df['订单日期'] <= conf_end)
    df_period = df.loc[mask].copy()

    # 定义“同一种商品”：通常指 商品名称 + 供应商 + 单价 均相同
    # 如果仅以名称为准，可改为 unique_items = df_period['商品名称'].nunique()
    unique_items_count = df_period.drop_duplicates(subset=['商品名称', '供应商', '单价（元）']).shape[0]

    # 动态提取品类
    sample_cats = df_period['商品名称'].str.split().str[0].unique()[:3]
    cat_str = "、".join(filter(lambda x: x not in ['nan', 'None', ''], sample_cats))

    # 5. 构造结果表格
    metrics_data = [
        {"维度": "全局概览", "指标": "采购企业总数", "数值": total_purchasers, "单位": "家", "关联信息": "---"},
        {"维度": "全局概览", "指标": "最高订单量企业", "数值": max_orders_val, "单位": "笔",
         "关联信息": max_orders_ent},
        {"维度": "全局概览", "指标": "最高采购金额企业", "数值": round(max_money_val, 2), "单位": "元",
         "关联信息": max_money_ent},
        {"维度": "区间概览", "指标": f"{month_label}交易商品种类", "数值": unique_items_count, "单位": "种",
         "关联信息": f"涵盖 {cat_str} 等品类"}
    ]
    metrics_df = pd.DataFrame(metrics_data)

    # 6. 打印控制台话术
    print("-" * 60)
    print(f"【报告话术参考】:")
    print(f"{month_label}产生交易商品 {unique_items_count} 种，涵盖 {cat_str} 等多品类。")
    print("-" * 60)

    # 7. 写回 Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for sn, s_df in all_sheets.items():
            if sn != "核心概览指标":
                s_df.to_excel(writer, sheet_name=sn, index=False)
        metrics_df.to_excel(writer, sheet_name="核心概览指标", index=False)

    print(f"核心指标提取完成。已将统计口径修正为：相同商品去重统计。")


if __name__ == "__main__":
    run_core_metrics()