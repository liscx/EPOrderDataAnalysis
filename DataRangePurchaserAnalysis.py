import pandas as pd
import os

# ================= 配置区域 =================
START_DATE = '2026-02-01'
END_DATE = '2026-02-28'
INPUT_FILE = '清洗结果_汇总.xlsx'


# ===========================================

def run_top10_analysis():
    if not os.path.exists(INPUT_FILE):
        print(f"错误：找不到文件 {INPUT_FILE}")
        return

    # 1. 加载所有工作表（为了最后完整写回）
    xls = pd.ExcelFile(INPUT_FILE)
    all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

    # 2. 提取并清洗数据
    df = all_sheets['清洗后数据'].copy()
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')

    # 筛选指定时间区间
    mask = (df['订单日期'] >= pd.to_datetime(START_DATE)) & (df['订单日期'] <= pd.to_datetime(END_DATE))
    df_period = df.loc[mask].copy()

    # 处理合并单元格空值
    df_period['订单号'] = df_period['订单号'].ffill()
    df_period['采购企业'] = df_period['采购企业'].ffill()
    df_period['专区名称'] = df_period['专区名称'].ffill()

    # 增加“省市”列：提取专区名称的前两个字
    df_period['省市'] = df_period['专区名称'].str[:2]

    # 3. 基础汇总（按企业聚合）
    ent_summary = df_period.groupby('采购企业').agg({
        '专区名称': lambda x: " / ".join(x.dropna().unique()),
        '省市': 'first',
        '订单号': 'nunique',  # 订单数量
        '数量': 'sum',  # 商品数量
        '订单金额（元）': 'sum'  # 订单金额
    }).reset_index()

    ent_summary.columns = ['采购企业', '专区名称', '省市', '订单数量', '商品数量', '订单金额（元）']

    # 4. 生成 Top 10 表格

    # --- 榜单1：按订单数量降序 ---
    top10_by_count = ent_summary.sort_values(by='订单数量', ascending=False).head(10).copy()
    top10_by_count.insert(0, '序号', range(1, len(top10_by_count) + 1))
    # 调整列顺序
    cols_count = ['序号', '采购企业', '专区名称', '省市', '订单数量', '商品数量', '订单金额（元）']
    top10_by_count = top10_by_count[cols_count]

    # --- 榜单2：按订单金额降序 ---
    top10_by_money = ent_summary.sort_values(by='订单金额（元）', ascending=False).head(10).copy()
    top10_by_money.insert(0, '序号', range(1, len(top10_by_money) + 1))
    # 调整列顺序
    cols_money = ['序号', '采购企业', '专区名称', '省市', '订单金额（元）', '订单数量', '商品数量']
    top10_by_money = top10_by_money[cols_money]

    # 5. 写回 Excel (保留原工作表，追加新排行页)
    with pd.ExcelWriter(INPUT_FILE, engine='openpyxl') as writer:
        for name, sheet_df in all_sheets.items():
            sheet_df.to_excel(writer, sheet_name=name, index=False)

        top10_by_count.to_excel(writer, sheet_name='采购企业订单量TOP10', index=False)
        top10_by_money.to_excel(writer, sheet_name='采购企业交易额TOP10', index=False)

    print(f"分析完成！已在 {INPUT_FILE} 中生成两个 Top10 排行榜。")
    print("\n--- 订单量 Top 1 企业 ---")
    print(top10_by_count.iloc[0][['采购企业', '订单数量']])
    print("\n--- 交易额 Top 1 企业 ---")
    print(top10_by_money.iloc[0][['采购企业', '订单金额（元）']])


if __name__ == "__main__":
    run_top10_analysis()
