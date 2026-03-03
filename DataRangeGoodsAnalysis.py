import pandas as pd
import os

# ================= 配置区域 =================
START_DATE = '2026-02-01'
END_DATE = '2026-02-28'
INPUT_FILE = '清洗结果_汇总.xlsx'


# ===========================================

def run_product_analysis():
    if not os.path.exists(INPUT_FILE):
        print(f"错误：找不到文件 {INPUT_FILE}")
        return

    # 1. 加载所有数据以备写回
    xls = pd.ExcelFile(INPUT_FILE)
    all_sheets = {name: xls.parse(name) for name in xls.sheet_names}
    df = all_sheets['清洗后数据'].copy()

    # 2. 预处理与区间筛选
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
    mask = (df['订单日期'] >= pd.to_datetime(START_DATE)) & (df['订单日期'] <= pd.to_datetime(END_DATE))
    df_p = df.loc[mask].copy()

    # 处理合并单元格产生的空值，确保商品行信息完整
    fill_cols = ['订单号', '供应商', '专区名称']
    df_p[fill_cols] = df_p[fill_cols].ffill()

    if df_p.empty:
        print("指定区间内无交易数据。")
        return

    # 3. 控制台概览统计
    total_items = int(df_p['数量'].sum())
    # 简单提取品类：取商品名称去重后的前3个作为代表
    sample_categories = df_p['商品名称'].str.split().str[0].unique()[:3]
    cat_str = "、".join(filter(None, sample_categories))

    # 统计月份交易商品件数（以1月为例，或根据区间动态展示）
    month_stats = df_p[df_p['订单日期'].dt.month == 1]
    jan_items = int(month_stats['数量'].sum()) if not month_stats.empty else 0

    print(f"--- 商品数据分析概览 ---")
    print(f"1月产生交易商品 {jan_items} 件，涵盖 {cat_str} 等多品类。")
    print(f"全时间段共计产生交易商品 {total_items} 件。")
    print("-" * 50)

    # 4. 商品维度聚合
    # 注意：商品名称+供应商+单价 作为唯一商品标识，防止同名不同价或不同供应商的商品混淆
    prod_grouped = df_p.groupby(['商品名称', '供应商', '单价（元）']).agg(
        销售数量=('数量', 'sum'),
        销售总额=('订单金额（元）', 'sum'),
        专区名称=('专区名称', lambda x: "、".join(x.dropna().unique()))
    ).reset_index()

    # 5. 生成 Top 10 榜单

    # --- 榜单1：销售数量 TOP 10 (按数量降序) ---
    top10_qty = prod_grouped.sort_values(by='销售数量', ascending=False).head(10).copy()
    top10_qty.insert(0, '序号', range(1, len(top10_qty) + 1))
    # 调整列顺序
    cols_qty = ['序号', '商品名称', '供应商', '销售数量', '单价（元）', '销售总额', '专区名称']
    # 这里的“单价（元）”即原表中的“单价（元）”，“销售总额”即聚合后的总额
    top10_qty = top10_qty.reindex(columns=cols_qty)
    top10_qty.rename(columns={'单价（元）': '单价\n（元）', '销售总额': '销售总额\n（元）'}, inplace=True)

    # --- 榜单2：销售金额 TOP 10 (按金额降序) ---
    top10_money = prod_grouped.sort_values(by='销售总额', ascending=False).head(10).copy()
    top10_money.insert(0, '序号', range(1, len(top10_money) + 1))
    # 调整列顺序
    cols_money = ['序号', '商品名称', '供应商', '销售总额', '销售数量', '单价（元）', '专区名称']
    top10_money = top10_money.reindex(columns=cols_money)
    top10_money.rename(columns={'单价（元）': '单价\n（元）', '销售总额': '销售总额\n（元）', '专区名称': '专区'},
                       inplace=True)

    # 6. 写回 Excel
    with pd.ExcelWriter(INPUT_FILE, engine='openpyxl') as writer:
        for name, sheet_df in all_sheets.items():
            sheet_df.to_excel(writer, sheet_name=name, index=False)

        top10_qty.to_excel(writer, sheet_name='商品销售数量TOP10', index=False)
        top10_money.to_excel(writer, sheet_name='商品销售金额TOP10', index=False)

    print(f"分析完成！已在 {INPUT_FILE} 中生成商品 Top10 排行榜。")


if __name__ == "__main__":
    run_product_analysis()