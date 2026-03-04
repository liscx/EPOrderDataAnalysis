import pandas as pd
import yaml
import os


def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_product_analysis():
    config = load_config()
    file_path = config['file_config']['output_file']
    start_date = config['analysis_period']['start_date']
    end_date = config['analysis_period']['end_date']

    if not os.path.exists(file_path):
        return

    # 1. 加载数据
    xls = pd.ExcelFile(file_path)
    all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

    # 优先从切片数据读取，确保时间区间绝对统一
    df_p = all_sheets.get('指定区间数据', pd.DataFrame())
    if df_p.empty:
        df = all_sheets['清洗后数据'].copy()
        df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
        mask = (df['订单日期'] >= pd.to_datetime(start_date)) & (df['订单日期'] <= pd.to_datetime(end_date))
        df_p = df.loc[mask].copy()

    # 2. 【数据对齐校准】
    # 补齐关键列，防止聚合丢失
    for col in ['订单号', '供应商', '专区名称', '商品名称', '单价（元）']:
        if col in df_p.columns:
            df_p[col] = df_p[col].ffill()

    # 3. 【核心修正】：销售金额 = 每一行的 (数量 * 当前行单价)
    # 这样可以处理同一商品在不同订单中单价略有波动的情况
    df_p['销售金额_计算'] = pd.to_numeric(df_p['数量'], errors='coerce') * pd.to_numeric(df_p['单价（元）'],
                                                                                         errors='coerce')

    # 4. 聚合统计
    # 注意：这里去掉了“单价”，因为同一商品单价可能波动。我们取平均单价展示，但按总数量聚合。
    prod_grouped = df_p.groupby(['商品名称', '供应商']).agg(
        销售数量=('数量', 'sum'),
        销售总额_计算=('销售金额_计算', 'sum'),
        平均单价=('单价（元）', 'mean'),  # 仅用于展示
        专区名称=('专区名称', lambda x: "、".join(x.dropna().astype(str).unique()))
    ).reset_index()

    # 5. 【排名决胜逻辑】：数量相同看金额，金额相同看名称（确保排名唯一稳定）
    # --- 榜单1：数量 TOP 10 ---
    top10_qty = prod_grouped.sort_values(
        by=['销售数量', '销售总额_计算', '商品名称'],
        ascending=[False, False, True]
    ).head(10).copy()

    top10_qty.insert(0, '序号', range(1, len(top10_qty) + 1))

    # --- 榜单2：金额 TOP 10 ---
    top10_money = prod_grouped.sort_values(
        by=['销售总额_计算', '销售数量', '商品名称'],
        ascending=[False, False, True]
    ).head(10).copy()

    top10_money.insert(0, '序号', range(1, len(top10_money) + 1))

    # 6. 写回 Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for name, sheet_df in all_sheets.items():
            if name not in ['商品销售数量TOP10', '商品销售金额TOP10']:
                sheet_df.to_excel(writer, sheet_name=name, index=False)
        top10_qty.to_excel(writer, sheet_name='商品销售数量TOP10', index=False)
        top10_money.to_excel(writer, sheet_name='商品销售金额TOP10', index=False)

    print("商品分析修正完成：已加入‘数量+金额’双重排序决胜逻辑。")