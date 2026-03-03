import pandas as pd
import yaml
import os


def load_config():
    """加载YAML配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_product_analysis():
    # 1. 加载配置
    config = load_config()
    file_path = config['file_config']['output_file']
    start_date = config['analysis_period']['start_date']
    end_date = config['analysis_period']['end_date']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 2. 确定数据源
    print(f"--- 正在执行：商品交易 Top10 分析 ---")
    xls = pd.ExcelFile(file_path)
    all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

    if '指定区间数据' in all_sheets:
        print("检测到 '指定区间数据' Sheet，直接读取...")
        df_p = all_sheets['指定区间数据'].copy()
    else:
        print("未检测到预切片数据，正在从 '清洗后数据' 中筛选日期...")
        df = all_sheets['清洗后数据'].copy()
        df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
        mask = (df['订单日期'] >= pd.to_datetime(start_date)) & \
               (df['订单日期'] <= pd.to_datetime(end_date))
        df_p = df.loc[mask].copy()

    if df_p.empty:
        print(f"警告：在指定区间 {start_date} 至 {end_date} 内无交易数据。")
        return

    # 3. 预处理：处理合并单元格产生的空值，确保商品行信息完整
    fill_cols = ['订单号', '供应商', '专区名称']
    for col in fill_cols:
        if col in df_p.columns:
            df_p[col] = df_p[col].ffill()

    # 4. 控制台概览统计
    total_items = int(df_p['数量'].sum()) if '数量' in df_p.columns else 0
    # 提取品类：取商品名称去重后的前3个作为代表
    sample_categories = df_p['商品名称'].astype(str).str.split().str[0].unique()[:3]
    cat_str = "、".join(filter(None, sample_categories))

    # 动态获取月份统计（以区间内第一个月为例）
    first_month = pd.to_datetime(start_date).month
    month_stats = df_p[pd.to_datetime(df_p['订单日期']).dt.month == first_month]
    m_items = int(month_stats['数量'].sum()) if not month_stats.empty else 0

    print(f"--- 商品数据分析概览 ---")
    print(f"{first_month}月产生交易商品 {m_items} 件，涵盖 {cat_str} 等多品类。")
    print(f"全时间段共计产生交易商品 {total_items} 件。")
    print("-" * 50)

    # 5. 商品维度聚合
    # 使用名称+供应商+单价作为唯一标识，防止不同规格或来源的商品混淆
    prod_grouped = df_p.groupby(['商品名称', '供应商', '单价（元）']).agg(
        销售数量=('数量', 'sum'),
        销售总额=('订单金额（元）', 'sum'),
        专区名称=('专区名称', lambda x: "、".join(x.dropna().astype(str).unique()))
    ).reset_index()

    # 6. 生成 Top 10 榜单

    # --- 榜单1：销售数量 TOP 10 ---
    top10_qty = prod_grouped.sort_values(by='销售数量', ascending=False).head(10).copy()
    top10_qty.insert(0, '序号', range(1, len(top10_qty) + 1))
    cols_qty = ['序号', '商品名称', '供应商', '销售数量', '单价（元）', '销售总额', '专区名称']
    top10_qty = top10_qty.reindex(columns=cols_qty)
    top10_qty.rename(columns={'单价（元）': '单价\n（元）', '销售总额': '销售总额\n（元）'}, inplace=True)

    # --- 榜单2：销售金额 TOP 10 ---
    top10_money = prod_grouped.sort_values(by='销售总额', ascending=False).head(10).copy()
    top10_money.insert(0, '序号', range(1, len(top10_money) + 1))
    cols_money = ['序号', '商品名称', '供应商', '销售总额', '销售数量', '单价（元）', '专区名称']
    top10_money = top10_money.reindex(columns=cols_money)
    top10_money.rename(columns={'单价（元）': '单价\n（元）', '销售总额': '销售总额\n（元）', '专区名称': '专区'},
                       inplace=True)

    # 7. 写回 Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for name, sheet_df in all_sheets.items():
            # 避开旧的商品排行 Sheet
            if name not in ['商品销售数量TOP10', '商品销售金额TOP10']:
                sheet_df.to_excel(writer, sheet_name=name, index=False)

        top10_qty.to_excel(writer, sheet_name='商品销售数量TOP10', index=False)
        top10_money.to_excel(writer, sheet_name='商品销售金额TOP10', index=False)

    print(f"分析完成！结果已同步至 {file_path}。")


if __name__ == "__main__":
    run_product_analysis()